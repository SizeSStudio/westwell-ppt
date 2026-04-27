import os
import tempfile
import unittest

from scripts.pptx_builder import WestwellPPT


SKILL_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
TEMPLATE = os.path.join(SKILL_DIR, 'assets', 'PPTTemplate.potx')


class VisualRefactorSmokeTest(unittest.TestCase):
    def test_new_visual_layouts_generate_pptx(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, 'visual-refactor-smoke.pptx')
            ppt = WestwellPPT(template=TEMPLATE, output=out)

            ppt.hero(
                title='AI Operator\n重塑运营边界',
                kicker='WESTWELL STRATEGY',
                lead='从单点自动化走向运营层智能。',
                meta_left='Visual grammar · Hero',
                meta_right='01 / 08',
            )
            ppt.big_numbers(
                title='三项指标说明方案进入可规模复制阶段',
                metrics=[
                    ('95%', '调度采纳率', '一线作业建议被稳定采纳'),
                    ('<1s', '响应时间', '关键调度路径实时返回'),
                    ('3x', '复制效率', '标准化实施周期显著压缩'),
                ],
                kicker='DATA EDITORIAL',
                lead='大数字先建立记忆点，解释文字只承担证据角色。',
            )
            ppt.pipeline(
                title='四步推进让试点从演示走向运营闭环',
                steps=[
                    {'title': '识别场景', 'body': '锁定高频、高损耗流程'},
                    {'title': '接入系统', 'body': '打通调度与执行数据'},
                    {'title': '人机共驾', 'body': '保留人工复核与兜底'},
                    {'title': '规模复制', 'body': '沉淀标准接口与指标'},
                ],
                kicker='SOFT SYSTEMS',
            )
            ppt.rowlines(
                title='资产优先级决定页面是否像真实方案',
                rows=[
                    ('Logo', '客户与 Westwell 标识', '必须确认'),
                    ('产品图', '车辆、系统、现场照片', '优先使用真实素材'),
                    ('UI 截图', '平台界面与运行态', '数字产品必备'),
                ],
            )
            ppt.quote_editorial(
                text='真正的智能调度，不是替代某个按钮，\n而是重新定义运营层的判断节奏。',
                attribution='Westwell visual grammar smoke test',
                kicker='EXECUTIVE MINIMAL',
            )
            ppt.lead_image(
                title='图片主导页必须保留顶部关键信息',
                lead='右侧图片采用标准比例和 top-fit 规则；缺素材时显示明确 placeholder。',
                img_path=None,
                kicker='IMAGE-LED',
            )
            ppt.image_grid(
                title='多素材页用统一比例建立秩序',
                images=[
                    {'path': None, 'label': 'Port scene'},
                    {'path': None, 'label': 'Dispatch UI'},
                    {'path': None, 'label': 'Vehicle detail'},
                    {'path': None, 'label': 'Energy node'},
                ],
                kicker='ASSET GRID',
            )
            ppt.bullets(
                '旧 API 仍可携带新版页眉页脚参数',
                ['保留现有调用方式', '新增参数不破坏旧代码', '输出仍为原生 PPTX'],
                meta_left='Legacy API',
                meta_right='08 / 08',
                foot_left='Smoke test',
                foot_right='Westwell PPT',
            )

            saved = ppt.save()

            self.assertTrue(os.path.exists(saved))
            self.assertGreater(os.path.getsize(saved), 0)
            self.assertEqual(len(ppt.prs.slides), 8)


if __name__ == '__main__':
    unittest.main()
