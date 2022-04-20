from configparser import ConfigParser
import numpy as np
import win32com.client as win32
from shapely.geometry import Point, LineString, Polygon
from shapely.affinity import rotate, scale
import Material


# cad 文件设置
acad = win32.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

# 读取配置文件
CONFIGFILE = 'xtract_config.ini'
config = ConfigParser()
config.read(CONFIGFILE)


def area_cal(shape_type, shape_param):
    """
    计算几何形状面积
    :param shape_type: 矩形 'rec'，圆形 'circle'
    :param shape_param: 数据格式列表，长宽[w, h]，直径[d]
    :return: 面积
    """
    if shape_type == 'rec':
        return shape_param[0] * shape_param[1]
    elif shape_type == 'circle':
        return np.pi * shape_param[0] ** 2 / 4
    else:
        raise Exception


def rec_pt(rec):
    pts = [
        [-rec[0] / 2, -rec[1] / 2],
        [-rec[0] / 2, rec[1] / 2],
        [rec[0] / 2, rec[0] / 2],
        [rec[0] / 2, -rec[1] / 2],
        [-rec[0] / 2, -rec[1] / 2]
    ]
    return pts


def pt_to_line_str(pts):
    str_list = ['Begin_Line'] + [','.join([str(j) for j in i]) for i in pts] + ['End_Line']
    return '\n'.join(str_list)


def pt_to_arc_str(pts):
    str_list = ['Begin_Arc'] + [','.join([str(j) for j in i]) for i in pts] + ['End_Arc']
    return '\n'.join(str_list)


def rebar_from_outline(outline_pts, space):
    outline_pts = np.array(outline_pts)
    rebar_pts = []
    for num, out_pt in enumerate(outline_pts[:-1]):
        line = [out_pt, outline_pts[num + 1]]
        length = np.linalg.norm(line[1] - line[0])
        pt_num = int(length // space) + 1
        for pt in range(pt_num):
            pt_id = line[0] + (line[1] - line[0]) * pt / pt_num
            rebar_pts.append(pt_id)
    return rebar_pts


def pts_to_rebar_str(pts, rebar_area, material, prestress=0):
    rebar_str = []
    for pt in pts:
        rebar_str.append(f'{pt[0]}, {pt[1]}, {rebar_area}, {prestress}, {material}')
    return '\n'.join(rebar_str)


class CadSection:
    def __init__(self, name, slt):
        self.name = name
        self.slt = slt
        self.center = [0, 0]
        self.mesh_size = 0
        self.shapes = []
        self.rebars = []

        self.members = {
            'outline': [],
            'rebar': []
        }

        # 统计 cad 几何信息，内外轮廓为多段线（polyline），钢筋为圆形（circle）
        for i in self.slt:
            if i.objectname == 'AcDbPolyline':
                self.members['outline'].append(i)
            elif i.objectname == 'AcDbCircle':
                self.members['rebar'].append(i)

        self.members['outline'].sort(key=lambda a: a.area, reverse=True)

        # 提取多段线数据
        outline_inf = []
        for out_l in self.members['outline']:
            outline_inf.append([])
            pts = np.array([round(i, 2) for i in out_l.coordinates]).reshape((-1, 2))
            for i, pt in enumerate(pts):
                i_inf = []
                bulge = out_l.getbulge(i)
                if bulge == 0:
                    i_inf.extend(['line', pt])
                else:
                    pt_next = pts[i + 1] if i < len(pts) - 1 else pts[0]
                    pt_center = (pt + pt_next) / 2
                    i_arc_half = scale(LineString([pt_center, pt]), xfact=bulge, yfact=bulge, origin=tuple(pt_center))
                    i_pt_arc = rotate(i_arc_half, 90, origin=tuple(pt_center)).coords[1]
                    i_inf.extend(['arc', pt, i_pt_arc])
                if i == len(pts) - 1:
                    i_inf.append(pts[0])
                outline_inf[-1].append(i_inf)

        # 整理截面整体信息
        self.section = {
            'box': [],
            'outline': [],
            'member': [],
            'rebar': []
        }
        for num, out_l in enumerate(outline_inf):
            pts = np.array([])
            for i in out_l:
                pts = np.append(pts, i[1:])
            i_box = Polygon(pts.reshape(-1, 2))
            if num == 0:
                self.center = np.array(i_box.centroid.coords[0])
                self.mesh_size = min(
                    i_box.bounds[2] - i_box.bounds[0],
                    i_box.bounds[3] - i_box.bounds[1]
                ) / 20
            elif not self.section['box'][0].contains(i_box.centroid):
                print('该轮廓不在范围内！')
                raise Exception
            self.section['box'].append(i_box)
            self.section['outline'].append(out_l)
            self.section['member'].append(self.members['outline'][num])

        for r in self.members['rebar']:
            if self.section['box'][0].contains(Point(r.center)):
                self.section['rebar'].append(r)


class XtractSection:
    def __init__(self, name, boundary):
        self.name = name
        self.boundary = boundary
        self.shapes = []
        self.rebars = []
        self.loadings = []

    def add_shape(self, material, mesh_size, line_and_arc):
        self.shapes.append(
            config['commands'].get('shape').format(
                material=material, mesh_size=mesh_size, line_and_arc=line_and_arc)
        )

    def add_rebar(self, rebar_part):
        self.rebars.append(rebar_part)

    def add_mc(self, name, const_axial, const_mxx, const_myy, inc, direction):
        """
        添加弯矩、曲率分析工况
        :param name: 荷载工况名称
        :param const_axial: 轴力，单位 N，压为负拉为正
        :param const_mxx: x弯矩，单位 N-m
        :param const_myy: y弯矩，单位 N-m
        :param inc: 增长项，可选为 IncAxial、IncMxx、IncMyy
        :param direction: 布尔值，True 代表正，False 代表负，轴力受压为负
        :return:
        """
        # 轴力为 4448，弯矩为 113
        if inc == 'IncAxial':
            direction_value = 4448
        elif inc in ['IncMxx', 'IncMyy']:
            direction_value = 113
        else:
            raise Exception

        direction_value = direction_value if direction else -direction_value

        self.loadings.append(
            config['commands'].get('moment_curvature').format(
                name=name,
                const_axial=const_axial,
                const_mxx=const_mxx,
                const_myy=const_myy,
                inc=inc,
                direction=direction_value
            )
        )

    def add_pm(self, name, full_or_half, angle):
        self.loadings.append(
            config['commands'].get('pm_interaction').format(
                name=name,
                full_or_half=full_or_half,
                angle=angle
            )
        )

    def add_mm(self, name, load):
        self.loadings.append(
            config['commands'].get('capacity_orbit').format(
                name=name,
                load=load
            )
        )

    def hollow_rectangle(self, out_rec, in_rec, cover, rebar_1, rebar_2, space_1, space_2, mat):
        """
        空心矩形截面
        :param mat: 混凝土、钢筋材料，['C50', 'HRB400']
        :param out_rec: 外轮廓尺寸 [w, h]，单位 mm
        :param in_rec: 内轮廓尺寸 [w, h]，单位 mm
        :param cover: 保护层厚度，单位mm
        :param rebar_1: 外圈钢筋直径，单位mm
        :param rebar_2: 内圈钢筋直径，单位mm
        :param space_1: 外圈钢筋间距，单位mm
        :param space_2: 内圈钢筋间距，单位mm
        :return:
        """
        mesh_size = min(out_rec[0] - in_rec[0], out_rec[1] - in_rec[1]) / 5

        line_str_1 = pt_to_line_str(
            rec_pt(out_rec)
        )
        line_str_2 = pt_to_line_str(
            rec_pt(in_rec)
        )
        self.add_shape(mat[0], mesh_size, line_str_1)
        self.add_shape('Delete', mesh_size, line_str_2)

        rebar_line_1 = rec_pt([i - cover * 2 - rebar_1 for i in out_rec])
        rebar_line_2 = rec_pt([i + cover * 2 + rebar_2 for i in in_rec])
        rebar_pts_1 = rebar_from_outline(rebar_line_1, space_1)
        rebar_pts_2 = rebar_from_outline(rebar_line_2, space_2)
        rebar_str_1 = pts_to_rebar_str(rebar_pts_1, area_cal('circle', [rebar_1]), mat[1]) + '\n'
        rebar_str_2 = pts_to_rebar_str(rebar_pts_2, area_cal('circle', [rebar_2]), mat[1])
        self.add_rebar(rebar_str_1)
        self.add_rebar(rebar_str_2)

    def section_from_cad(self, mesh_size=0):

        # 框选截面图形
        slt = doc.SelectionSets.Add(self.name)
        slt.SelectOnScreen()
        cad_sec = CadSection(self.name, slt)

        # 确定网格尺寸
        mesh_size = cad_sec.mesh_size if mesh_size == 0 else mesh_size

        # 生成形状和钢筋
        for num, out_l in enumerate(cad_sec.section['outline']):
            all_line_str = []
            for line in out_l:
                if line[0] == 'line':
                    pts = [np.array(i) - cad_sec.center for i in line[1:]]
                    all_line_str.append(pt_to_line_str(pts))

                elif line[0] == 'arc':
                    pts = [np.array(i) - cad_sec.center for i in line[1:]]
                    all_line_str.append(pt_to_arc_str(pts))
                else:
                    raise Exception
            self.add_shape(
                cad_sec.section['member'][num].layer,
                mesh_size,
                '\n'.join(all_line_str)
            )

        rebar_str = []
        for r in cad_sec.section['rebar']:
            r_pt = np.array(r.center[:2]) - cad_sec.center
            rebar_str.append(f'{r_pt[0]}, {r_pt[1]}, {r.area}, 0, {r.layer}')
        self.add_rebar('\n'.join(rebar_str))


class XtractXpj:
    def __init__(self, name):
        self.mat_abbreviation = {
            'c1': self.unconfined_concrete,
            'r1': self.bilinear_rebar
        }

        self.overall = config['commands'].get('global').format(name=name)
        self.materials = []
        self.sections = []

    def xpj_command(self):
        """
        完整项目命令流
        :return: 命令流
        """
        commands = [self.overall] + self.materials + self.sections
        return '\n'.join(commands)

    def unconfined_concrete(self, name, strength):
        self.materials.append(
            config['commands'].get('unconfined_concrete').format(name=name, strength=strength)
        )

    def bilinear_rebar(self, name, stress):
        self.materials.append(
            config['commands'].get('bilinear_rebar').format(name=name, stress=stress)
        )

    def add_materials(self, materials):
        """
        一次性添加多种材料
        :param materials: [['c1', 'C50', 22.4], ['r1', 'HRB400', 330]]
        """
        for mat in materials:
            self.mat_abbreviation[mat[0]](*mat[1:])

    def add_section(self, section):
        """截面命令流"""
        shapes_str = '\n'.join(section.shapes)
        rebars_str = '\n'.join(section.rebars)
        loadings_str = '\n'.join(section.loadings)

        self.sections.append(
            config['commands'].get('section').format(
                name=section.name,
                boundary=section.boundary,
                shape_part=shapes_str,
                rebar_part=rebars_str,
                loading_part=loadings_str
            )
        )


if __name__ == '__main__':
    # 创建项目
    c1 = Material.Concrete.C40
    r1 = Material.Rebar.HRB500

    mats = [['c1', 'C40', c1.fcd], ['r1', 'HRB500', r1.fsd]]
    test_xpj = XtractXpj('NupPier')
    test_xpj.add_materials(mats)

    # 空心矩形截面测试
    # test_hollow_sec = XtractSection('hollow_sec_1', 5000)
    # test_hollow_sec.hollow_rectangle([3000, 3000], [2000, 2000], 75, 45, 45, 100, 100, [i[1] for i in mats])
    # test_hollow_sec.add_mc('test_mc', -1, 1, 0, 'IncMxx', True)
    # test_hollow_sec.add_mm('test_mm', -16679e3)
    # test_xpj.add_section(test_hollow_sec)

    # cad 截面测试
    test_cad_sec = XtractSection('cad_sec_2', 3000)
    test_cad_sec.section_from_cad(mesh_size=100)
    # test_cad_sec.add_mc('test_mc', -1, 1, 0, 'IncMxx', True)
    test_cad_sec.add_mm('100kn_mm', -1000e3)
    test_xpj.add_section(test_cad_sec)

    # 写入文件
    with open('NupPier_1.8_125.xpj', 'w') as f:
        f.write(test_xpj.xpj_command())



