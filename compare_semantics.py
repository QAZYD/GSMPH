import pandas as pd
from sentence_transformers import SentenceTransformer, util

print("1. Loading SBERT Multilingual model...")
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

excel_file = 'GSMPH_Cleaned_Terms.xlsx'
target_sheet = 'Lookup Table (Cleaned)'
print(f"2. Reading sheet '{target_sheet}' from {excel_file}...")

df = pd.read_excel(excel_file, sheet_name=target_sheet)

col_orig = 'Term'
col_loc = 'Previous Replacement' if 'Previous Replacement' in df.columns else 'Replacement Pool'

print(f"\n3. Comparing semantic similarity between: '{col_orig}' VS '{col_loc}'")

original_texts = df[col_orig].fillna('').astype(str).tolist()
localized_texts = df[col_loc].fillna('').astype(str).tolist()

print("4. Generating semantic embeddings...")
embeddings_orig = model.encode(original_texts, convert_to_tensor=True)
embeddings_loc = model.encode(localized_texts, convert_to_tensor=True)

print("5. Computing similarity scores...")
cosine_scores = util.cos_sim(embeddings_orig, embeddings_loc)

# Extract scores per row
scores = [cosine_scores[i][i].item() for i in range(len(df))]
df['semantic_similarity_score'] = scores

# Flag Outliers (< 0.70 threshold)
threshold = 0.70
df['needs_human_review'] = df['semantic_similarity_score'] < threshold

# Save output
output_file = 'semantic_evaluation_results.csv'
df.to_csv(output_file, index=False)

print(f"\nSUCCESS! Results saved to '{output_file}'.")
print(f"Total rows evaluated: {len(df)}")
print(f"Outliers flagged for human review (< {threshold}): {df['needs_human_review'].sum()}")