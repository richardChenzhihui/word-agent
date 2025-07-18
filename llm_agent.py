# llm_word_agent.py  — 轻量优化 & 稳定版（含表格 / 单元格自动解析）
import json, inspect, textwrap, re
from typing import Dict, List, Any, Optional

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell  # ← 新增 _Cell

# import docx_utils as du  # Word 操作函数集合
import docx_utils as du

# --------------------------  LLMConfig  -------------------------- #
class LLMConfig:
    def __init__(
            self,
            api_key: str,
            model: str = "gpt-3.5-turbo",
            base_url: Optional[str] = None,
            api_type: str = "openai",
            api_version: Optional[str] = None,
            temperature: float = 0.2,
            max_tokens: int = 2000,
            organization: Optional[str] = None,
    ):
        self.api_key, self.model = api_key, model
        self.base_url, self.api_type, self.api_version = base_url, api_type, api_version
        self.temperature, self.max_tokens, self.organization = temperature, max_tokens, organization


# =============================  Agent  ============================= #
class LLMWordAgent:
    def __init__(self, llm_config: Optional[LLMConfig] = None):
        self.llm_config = llm_config
        self.tools_meta = self._scan_tools()  # {name: (func, sig, intro)}
        self.utils_doc = self._build_tools_doc()
        self.runtime_ctx = {"last_table": None}  # ← 运行期上下文
        self._setup_llm_client()

    # ---------- 扫描 docx_utils ---------- #
    def _scan_tools(self) -> Dict[str, tuple]:
        tools = {}
        for name, fn in vars(du).items():
            if name.startswith("_") or not inspect.isfunction(fn):
                continue
            sig = inspect.signature(fn)
            doc_str = (fn.__doc__ or "").strip()
            intro = doc_str.splitlines()[0] if doc_str else ""
            tools[name] = (fn, sig, intro)
        return tools

    def _build_tools_doc(self) -> str:
        lines = ["【Word 编辑工具 - 自动生成】"]
        for name, (_, sig, intro) in self.tools_meta.items():
            lines.append(f"{name}{sig}  –  {intro}")
        return "\n".join(lines)

    # ---------- LLM 客户端 ---------- #
    def _setup_llm_client(self):
        if not self.llm_config:
            self.client = None
            return

        import openai

        params = {"api_key": self.llm_config.api_key}
        if self.llm_config.organization:
            params["organization"] = self.llm_config.organization
        if self.llm_config.base_url:
            if self.llm_config.api_type == "azure":
                params.update(
                    azure_endpoint=self.llm_config.base_url,
                    api_version=self.llm_config.api_version or "2024-02-15-preview",
                )
            else:
                params["base_url"] = self.llm_config.base_url
        self.client = openai.OpenAI(**params)

    # ---------- Prompt ---------- #
    def _build_prompt(self, structure: Dict[str, Any], user_inst: str) -> str:
        return textwrap.dedent(
            f"""
            你是专业 Word 编辑助手。请根据用户指令，输出 docx_utils 调用指令。
            【可用工具】
            {self.utils_doc}

            【当前文档结构】
            {json.dumps(structure, ensure_ascii=False, indent=2)}

            【用户指令】
            {user_inst}

            【输出要求】
            1. 仅输出 JSON 数组，每项包含 "action" 及对应参数，字段名与函数签名一致。
            2. 若函数首参为 doc（Document 对象）系统会自动注入，JSON 中不要提供 doc。
            3. Paragraph / Table 参数用 index（int）或 {{ "index": idx }} 表示。
            4. 如需单元格，可仅提供 {{ "row": r, "col": c }}，系统自动定位到最近创建的表格。
            5. 如果忘记提供 table 参数，系统默认使用最近一次 add_table 创建的表格。
            6. 不要输出除 JSON 之外的任何文本。
            """
        )

    # ---------- 调用 LLM ---------- #
    def call_llm(self, prompt: str) -> List[Dict[str, Any]]:
        if not self.client:  # 离线演示
            return [{"action": "add_heading", "text": "示例文档", "level": 1}]

        rsp = self.client.chat.completions.create(
            model=self.llm_config.model,
            messages=[
                {"role": "system", "content": "你是 Word 编辑助手，仅输出 JSON"},
                {"role": "user", "content": prompt},
            ],
            temperature=self.llm_config.temperature,
            max_tokens=self.llm_config.max_tokens,
            response_format={"type": "json_object"},
        )
        content = rsp.choices[0].message.content
        print("-----------------\n", content, "\n-----------------")

        # 解析 JSON
        try:
            data = json.loads(content)
            return data if isinstance(data, list) else data.get("commands", [])
        except json.JSONDecodeError:
            m = re.search(r"\[.*\]", content, re.S)
            if m:
                return json.loads(m.group(0))
            raise ValueError("LLM 输出无法解析为 JSON")

    # ---------- 参数校验 & 转换 ---------- #
    def _validate_and_prepare(self, doc: Document, cmd: Dict[str, Any]):
        action = cmd.get("action")
        if action not in self.tools_meta:
            raise ValueError(f"未知 action: {action}")

        fn, sig, _ = self.tools_meta[action]
        kwargs = {k: v for k, v in cmd.items() if k != "action"}

        # -------- 始终注入 / 覆盖 doc -------- #
        for p in sig.parameters.values():
            if p.annotation is Document or p.name == "doc":
                kwargs[p.name] = doc

        # -------- 特殊处理 set_cell_text 的 row/col 参数 -------- #
        if action == "set_cell_text" and "row" in kwargs and "col" in kwargs:
            # 从 row/col 构造 cell 参数
            row = kwargs.pop("row")
            col = kwargs.pop("col")
            table_index = kwargs.pop("table_index", None)
            
            if table_index is not None:
                table = doc.tables[table_index]
            elif self.runtime_ctx.get("last_table"):
                table = self.runtime_ctx["last_table"]
            else:
                raise ValueError("无法确定表格：缺少 table_index 且没有最近创建的表格")
            
            kwargs["cell"] = table.cell(row, col)

        # -------- 将 index / dict 转对象 -------- #
        def to_para(v):
            if isinstance(v, int):
                return doc.paragraphs[v]
            if isinstance(v, dict):
                return doc.paragraphs[v.get("index", 0)]
            return v

        def to_table(v):
            if isinstance(v, int):
                return doc.tables[v]
            if isinstance(v, dict):
                return doc.tables[v.get("index", 0)]
            # v 为空 → 使用最近表格（若存在）
            if v is None and self.runtime_ctx.get("last_table"):
                return self.runtime_ctx["last_table"]
            return v

        def to_cell(v):
            if isinstance(v, _Cell):
                return v
            if isinstance(v, dict) and {"row", "col"} <= v.keys():
                tbl = (
                    doc.tables[v.get("table_index")]
                    if "table_index" in v
                    else to_table(None)
                )
                return tbl.cell(v["row"], v["col"])
            return v

        # 遍历函数参数，自动转换
        for p in sig.parameters.values():
            name_lower = p.name.lower()
            ann = p.annotation

            # 若参数缺失但需要 Table，对其使用最近表格
            if p.name not in kwargs and (
                    ann is Table or name_lower == "table"
            ) and self.runtime_ctx.get("last_table"):
                kwargs[p.name] = self.runtime_ctx["last_table"]

            if p.name not in kwargs:
                continue

            val = kwargs[p.name]

            if ann is Paragraph or (
                    ann is inspect._empty and "para" in name_lower
            ):
                kwargs[p.name] = to_para(val)

            elif ann is Table or (
                    ann is inspect._empty and "table" in name_lower
            ):
                kwargs[p.name] = to_table(val)

            elif ann is _Cell or (
                    ann is inspect._empty and "cell" in name_lower
            ):
                kwargs[p.name] = to_cell(val)

        # -------- 必填 / 多余检查 -------- #
        for p in sig.parameters.values():
            if p.default is p.empty and p.name not in kwargs:
                raise ValueError(f"缺少必需参数 {p.name} for {action}")
        extra = set(kwargs) - set(sig.parameters)
        if extra:
            raise ValueError(f"多余参数 {extra} for {action}")

        return fn, kwargs

    # ---------- 执行指令 ---------- #
    def execute_commands(self, doc: Document, cmds: List[Dict[str, Any]]):
        res = {"success": [], "errors": []}
        for idx, cmd in enumerate(cmds, 1):
            try:
                fn, kwargs = self._validate_and_prepare(doc, cmd)
                out = fn(**kwargs)  # 真正调用

                # 若执行 add_table，记录返回表格
                if cmd["action"] == "add_table" and isinstance(out, Table):
                    self.runtime_ctx["last_table"] = out

                res["success"].append(f"#{idx} {cmd['action']} OK")
            except Exception as e:
                res["errors"].append(f"#{idx} {cmd}: {e}")
        return res

    # ---------- 顶层接口 ---------- #
    def edit_document(self, docx_path: str, user_inst: str, out_path: str):
        try:
            doc = du.open_docx(docx_path)
            structure = du.export_docx_structure(doc)
            prompt = self._build_prompt(structure, user_inst)
            commands = self.call_llm(prompt)
            exec_res = self.execute_commands(doc, commands)
            du.save_docx(doc, out_path)
            return {
                "status": "success",
                "commands": commands,
                "execution_results": exec_res,
                "output": out_path,
            }
        except Exception as e:
            return {"status": "error", "error": str(e)}


# ------------------------------ 示例 ------------------------------ #
if __name__ == "__main__":
    """
    直接运行此文件可执行一次“多功能综合测试”，需：
    • demo.docx 作为输入
    • finance.png（或自行修改图片路径）位于同级目录
    • 配置自己的 OpenAI 兼容 key
    """
    import os

    os.environ["OPENAI_API_KEY"] = "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"

    cfg = LLMConfig(
        api_key=os.getenv("OPENAI_API_KEY"),
        base_url="https://openrouter.ai/api/v1",
        model="moonshotai/kimi-k2",
        temperature=0.7,
        # max_tokens=2048,
    )
    agent = LLMWordAgent(cfg if cfg.api_key else None)

    # user_inst = """
    # 请按以下顺序对文档进行多步骤编辑：
    #
    # 1) 在文档开头插入一级标题“示例文档”。
    # 2) 紧跟其后插入自我介绍段落：“大家好，我是大型语言模型驱动的 Word 智能助手，下面演示多功能编辑。”
    # 3) 在自我介绍段落后插入分页符。
    # 4) 在新页添加二级标题“香港金融发展展望”。
    # 5) 在标题后插入三段文字（你生成），讨论香港金融业未来 5 年的主要趋势。
    # 6) 然后插入一个无序列表，内容为：绿色金融、离岸人民币、金融科技。
    # 7) 接着插入一个 3×3 表格：
    #    - 第一行（标题行）：指标 | 2023 | 2024 （并将这一整行 3 个单元格合并）
    #    - 第二行：IPO数量 | 90 | 120
    #    - 第三行：融资总额 | 1200亿 | 1500亿
    #    - 表格样式设为 “Table Grid”。
    # 8) 把正文和表格中出现的“香港”全部替换为“H.K.”（保留格式）。
    # 9) 在表格后再插入分页符并插入图片 finance.png，宽 4 英寸。
    # 10) 在文档末尾插入一个有序列表，列出你认为香港金融可持续增长的 3 个关键词。
    # 11) 设置文档作者属性为 “LLM-Agent”。
    # 12) 把所有的文字替换成英文翻译，你（大模型）提供翻译
    # """
    #
    # user_inst = """
    # Please read the existing contract text in the current document and replace each occurrence of the old contract elements listed below with the corresponding new values. Do not modify any other clauses, numbering, or formatting—only perform precise text substitutions. Return **only** a JSON array of `docx_utils` commands.
    #
    # New contract elements to apply:
    # - Party A: Beijing ZhiXin Technology Co., Ltd.
    # - Party B: Shanghai Internet Network Co., Ltd.
    # - Signing Date: July 1, 2025
    # - Performance Location: Wangjing Street, Chaoyang District, Beijing
    # - Total Contract Amount: RMB 800,000.00
    #
    # And change all text into time new roman
    # """
    #
    # res = agent.edit_document(
    #     docx_path="demo.docx",
    #     user_inst=user_inst,
    #     out_path="demo_out.docx",
    # )
    # print(json.dumps(res, ensure_ascii=False, indent=2))
    #
    # user_inst = """
    # Please read the existing legal text in the current document and replace each occurrence of the old contract elements listed below with the corresponding new values. Do not modify any other clauses, numbering, or formatting—only perform precise text substitutions. Return **only** a JSON array of `docx_utils` commands.
    #
    # New contract elements to apply:
    # initial limited partner 的名字和信息如下：姓名：秦易，身份证地址：四川省广安市广安区龙台镇竹埝村3组12号。
    # input text should be the english translation, you should read the document and find all places need modification
    # """
    #
    # res = agent.edit_document(
    #     docx_path="./documents/1. Sierra Education LPF - Initial LPA - Execution Version.docx",
    #     user_inst=user_inst,
    #     out_path="./documents/1. Sierra Education LPF - Initial LPA - Execution Version_out.docx",
    # )
    # print(json.dumps(res, ensure_ascii=False, indent=2))

    user_inst = """
    写一个 陈致晖 AI 算法实习的交接文档
    交接包括两部分的技术

    dify 的新推理 发布在了tag news ollama， 使用ollama + qwen3 8/32b 模型推理，主要的prompt修改在：
    添加而输出结构限制：

    DO NOT include anything not the direct answer. Only output direct answer.

    direct answer (do not include this):

    同时在输出端剔除思考模型的<think/> </think> 思考内容，示例代码
    if '<think>' in company and '</think>' in company:
        # 找到</think>后的内容，通常是"Direct Answer:"或直接是答案
        think_end = company.find('</think>')
        answer_part = company[think_end + 8:].strip()  # 8是'</think>'的长度

        company = answer_part

    同时调整prompt规范模型输出格式


     gitlab 上 AI News Evaluation 加了 tag_news_ollama 的branch，更新了代码
     测试代码修改了api，可以直接运行发布了的dify tag news ollama 的测试


     后续微调工作：写清楚 推荐使用verl 字节火山架构做 llm 强化学习，架构代码位于服务器/home/aqumon/ai_news/local_llm_finetune，架构环境配置于conda/verl                   /home/aqumon/anaconda3/envs/verl
     随后附上你（大模型）生成的适合算法交接的verl 使用简短说明

     后续微调工作推荐的策略：选用8B模型，配合1-2k标好的数据，使用grpo强化学习，reward设置为标签的正确数量，基于verl进行微调，学习人类的标签偏好

     最后结束

     以上所有的文字需要你进行拓展，组织，生成一个有各级标题的交接文档, 
     最后把全部更新后的字体全部改为宋体

    """

    res = agent.edit_document(
        docx_path="documents/AQUMON.docx",
        user_inst=user_inst,
        out_path="documents/AQUMON_out.docx",
    )
    print(json.dumps(res, ensure_ascii=False, indent=2))
    #

    # 已经发现的bug
    # 1. 原本的docx 不能是空的否则无法读取
    # 2. add_list 有问题
    # 3. 律所文件出现发现执行修改后，word 没有变动的地方
