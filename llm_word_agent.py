# llm_word_agent.py  – 轻量优化 & 稳定版
import json, inspect, textwrap, re
from typing import Dict, List, Any, Optional
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import docx_utils as du                     # 工具库

# -------- LLMConfig（保持原设计） -------- #
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


# ===================== Agent ===================== #
class LLMWordAgent:
    def __init__(self, llm_config: Optional[LLMConfig] = None):
        self.llm_config   = llm_config
        self.tools_meta   = self._scan_tools()   # {name: (func, sig, intro)}
        self.utils_doc    = self._build_tools_doc()
        self._setup_llm_client()

    # ---------- 扫描 docx_utils ---------- #
    def _scan_tools(self) -> Dict[str, tuple]:
        tools = {}
        for name, fn in vars(du).items():
            if name.startswith('_') or not inspect.isfunction(fn):
                continue
            sig = inspect.signature(fn)
            doc_str   = (fn.__doc__ or '').strip()
            intro     = doc_str.splitlines()[0] if doc_str else ''
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
        return textwrap.dedent(f"""
        你是专业 Word 编辑助手。请根据用户指令，生成对文档的编辑操作指令。
        【可用工具】
        {self.utils_doc}

        【当前文档结构】
        {json.dumps(structure, ensure_ascii=False, indent=2)}

        【用户指令】
        {user_inst}

        【输出要求】
        1. 只输出 JSON 数组，每个对象必须包含 "action" 与相应参数，字段名与函数签名一致。
        2. 函数首参若为 doc（Document 对象）系统会自动注入，JSON 中不要提供 doc。
        3. 若需 Paragraph/Table，请提供其 index（int）或包含 index 字段的对象。
        4. 不要输出多余文字。
        
        如果忘记提供 table 参数，系统默认使用最近一次 add_table 创建的表格。  
        对于单元格可仅提供 {{"row": r, "col": c}}，系统会自动定位到当前表格。
        """)

    # ---------- 调用 LLM ---------- #
    def call_llm(self, prompt: str) -> List[Dict[str, Any]]:
        if not self.client:                        # 离线演示
            return [{"action": "add_heading", "text": "示例文档", "level": 1}]

        rsp = self.client.chat.completions.create(
            model=self.llm_config.model,
            messages=[
                {"role": "system", "content": "你是 Word 编辑助手，仅输出 JSON"},
                {"role": "user",   "content": prompt},
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
            m = re.search(r'\[.*\]', content, re.S)
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

        # 始终注入/覆盖 doc 参数
        for p in sig.parameters.values():
            if p.annotation is Document or p.name == "doc":
                kwargs[p.name] = doc

        # 将 index / dict 转成 Paragraph/Table
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
            return v

        for p in sig.parameters.values():
            if p.name not in kwargs:
                continue
            ann = p.annotation
            val = kwargs[p.name]
            # 根据注解或参数名关键字判断
            if ann is Paragraph or (ann is inspect._empty and "para" in p.name.lower()):
                kwargs[p.name] = to_para(val)
            elif ann is Table or (ann is inspect._empty and "table" in p.name.lower()):
                kwargs[p.name] = to_table(val)

        #  必填/多余检查
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
                fn(**kwargs)
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
            cmds   = self.call_llm(prompt)
            exec_res = self.execute_commands(doc, cmds)
            du.save_docx(doc, out_path)
            return {
                "status": "success",
                "commands": cmds,
                "execution_results": exec_res,
                "output": out_path,
            }
        except Exception as e:
            return {"status": "error", "error": str(e)}


# ---------------- 示例 ---------------- #
if __name__ == "__main__":
    import os
    os.environ["OPENAI_API_KEY"] = "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"
    # !! 请把自己的 Key 写到环境变量，示例仅供演示 !!
    os.environ.setdefault("OPENAI_API_KEY", "")
    cfg = LLMConfig(
        api_key=os.getenv("OPENAI_API_KEY", ""),
        base_url="https://openrouter.ai/api/v1",
        model="moonshotai/kimi-k2",
        temperature=0.1,
        max_tokens=1024,
    )

    # agent = LLMWordAgent(cfg if cfg.api_key else None)
    # result = agent.edit_document(
    #     docx_path="demo.docx",
    #     user_inst="在开头插入标题“示例文档”；在文档最后添加文字“猪猪”及你（大模型）对香港金融发展的展望，在文档最前面添加你（大模型）的自我介绍",
    #     out_path="demo_out.docx",
    # )
    # print(json.dumps(result, ensure_ascii=False, indent=2))

    agent = LLMWordAgent(cfg if cfg.api_key else None)

    # 2. 一条指令尽可能覆盖最多功能
    user_inst = """
    请按以下顺序对文档进行多步骤编辑：

    1) 在文档开头插入一级标题“示例文档”。  
    2) 紧跟其后插入自我介绍段落：“大家好，我是大型语言模型驱动的 Word 智能助手，下面演示多功能编辑。”  
    3) 在自我介绍段落后插入分页符。  
    4) 在新页添加二级标题“香港金融发展展望”。  
    5) 在标题后插入三段文字（你生成），讨论香港金融业未来 5 年的主要趋势。  
    6) 然后插入一个无序列表，内容为：绿色金融、离岸人民币、金融科技。  
    7) 接着插入一个 3×3 表格：  
       - 第一行（标题行）：指标 | 2023 | 2024 （并将这一整行 3 个单元格合并）  
       - 第二行：IPO数量 | 90 | 120  
       - 第三行：融资总额 | 1200亿 | 1500亿  
       - 表格样式设为 “Table Grid”。  
    8) 把正文和表格中出现的“香港”全部替换为“H.K.”（保留格式）。  
    9) 在表格后再插入分页符并插入图片 finance.png，宽 4 英寸。  
    10) 在文档末尾插入一个有序列表，列出你认为香港金融可持续增长的 3 个关键词。  
    11) 设置文档作者属性为 “LLM-Agent”。  
    """

    result = agent.edit_document(
        docx_path="demo.docx",  # 原始文件
        user_inst=user_inst,
        out_path="demo_out.docx"  # 输出文件
    )

    print(json.dumps(result, ensure_ascii=False, indent=2))

