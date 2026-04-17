import os
import pandas as pd
from tabulens import TableExtractor

# Map internal gateway env vars to OpenAI-compatible ones used by langchain_openai
if not os.getenv("OPENAI_BASE_URL") and os.getenv("ANTHROPIC_BASE_URL"):
    os.environ["OPENAI_BASE_URL"] = os.getenv("ANTHROPIC_BASE_URL", "")
if not os.getenv("OPENAI_API_KEY") and os.getenv("ANTHROPIC_API_KEY"):
    os.environ["OPENAI_API_KEY"] = os.getenv("ANTHROPIC_API_KEY", "")

extractor = TableExtractor(model_name="gpt:glm-4-32b")

tables = extractor.extract_tables("input/PCI_243_245.pdf")

os.makedirs("output/tabulen", exist_ok=True)

for i, table in enumerate(tables):
    if table is None:
        print(f"[warn] table #{i} is None, skipping")
        continue
    try:
        if hasattr(table, "to_pandas"):
            df = table.to_pandas()
        elif isinstance(table, pd.DataFrame):
            df = table
        else:
            print(f"[warn] table #{i} has unsupported type: {type(table)}")
            continue
    except Exception as e:
        print(f"[warn] table #{i} to_pandas failed: {e}")
        continue

    out_path = f"output/tabulen/table_{i}.csv"
    df.to_csv(out_path, index=False)
    print(f"[ok] wrote {out_path} (rows={len(df)})")
