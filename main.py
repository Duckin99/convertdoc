import os
from openai import OpenAI
from docx2md.pipeline import DocxPipeline
from docx2md.agents import VisionAgent, MarkdownAgent

if __name__ == "__main__":
    DOC_FILE = "sample_large_doc.docx"
    
    # Initialize the LLM client
    # client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    
    # 1. Initialize Pipeline
    pipeline = DocxPipeline(DOC_FILE, out_dir="output_structured")
    
    # 2. Attach Agents (Uncomment to enable End-to-End AI mode)
    # v_agent = VisionAgent(client)
    # md_agent = MarkdownAgent(client)
    # pipeline.attach_agents(vision_agent=v_agent, md_agent=md_agent)
    
    # 3. Run
    pipeline.execute()