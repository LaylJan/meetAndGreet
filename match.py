import pandas as pd

# === Read Excel file ===
df = pd.read_excel("Attendees.xlsx")

# Clean column names
df.columns = df.columns.str.strip()

# Identify key and preference columns
non_pref_cols = ["Name", "Lokal", "Gender"]
pref_cols = [col for col in df.columns if col not in non_pref_cols]

# Convert all preference columns to numeric (non-numbers → 0)
for col in pref_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# Compatibility function
def compatibility_score(p1, p2):
    scores1 = p1[pref_cols].values
    scores2 = p2[pref_cols].values
    diff = abs(scores1 - scores2)
    score = 100 - (diff.mean() * 20)
    return round(score, 2)

print("💞 TOP 3 OPPOSITE-GENDER MATCHES FOR EVERYONE 💞\n")

# Go through everyone and find their top 3 opposite-gender matches
for i, person in df.iterrows():
    name = person["Name"]
    lokal = person["Lokal"]
    gender = str(person["Gender"]).strip().lower()

    # Opposite gender filter
    if gender == "male":
        others = df[(df["Name"] != name) & (df["Gender"].str.lower() == "female")]
    elif gender == "female":
        others = df[(df["Name"] != name) & (df["Gender"].str.lower() == "male")]
    else:
        continue  # skip if gender is not male/female

    matches = []
    for j, other in others.iterrows():
        score = compatibility_score(person, other)
        matches.append((other["Name"], other["Lokal"], other["Gender"], score))

    # Sort by compatibility descending and get top 3
    top_matches = sorted(matches, key=lambda x: x[3], reverse=True)[:3]

    print(f"👤 {name} ({lokal}, {gender.title()}) — Top 3 Matches:")
    for match in top_matches:
        print(f"   ❤️ {match[0]} ({match[1]}) → Compatibility: {match[3]}%")
    print()

print("✅ Matching complete! Everyone has their Top 3 opposite-gender matches.")
