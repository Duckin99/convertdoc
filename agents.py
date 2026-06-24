import base64
from pathlib import Path
from openai import OpenAI

class VisionAgent:
    def __init__(self, client: OpenAI, model: str = "gpt-4o"):
        self.client = client
        self.model = model

    def read_img(self, img_path: Path) -> str:
        if not img_path.exists():
            return f"\n[Missing Image: {img_path.name}]\n"

        with open(img_path, "rb") as f:
            b64_img = base64.b64encode(f.read()).decode('utf-8')

        prompt = (
            "Extract content from this image for an AI database.\n"
            "If diagram/flowchart: Output Mermaid.js code.\n"
            "If table: Output Markdown table.\n"
            "If decorative: Reply exactly with \n"
            "Output ONLY the code or text."
        )

        try:
            res = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"}}]}]
            )
            return f"\n\n{res.choices[0].message.content.strip()}\n\n"
        except Exception as e:
            return f"\n[Vision Error: {str(e)}]\n"

class MarkdownAgent:
    def __init__(self, client: OpenAI, model: str = "gpt-4o"):
        self.client = client
        self.model = model

    def clean_markdown(self, raw_md: str) -> str:
        """Acts as an end-to-end formatter to ensure the final MD is perfect for RAG/AI."""
        prompt = (
            "You are a data pipeline agent. I will provide raw, scraped markdown from a Word document.\n"
            "Your job is to clean it up so it is perfectly optimized for another LLM or RAG system to read.\n"
            "- Fix any broken markdown tables.\n"
            "- Fix weird spacing or broken bulleted lists.\n"
            "- Remove redundant Table of Contents if they just clutter the file.\n"
            "- Maintain all Mermaid blocks and links.\n"
            "Output ONLY the cleaned markdown file, without any conversational text."
        )

        try:
            res = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": raw_md}
                ]
            )
            return res.choices[0].message.content.strip()
        except Exception as e:
            print(f"Agent Cleanup Error: {e}")
            return raw_md