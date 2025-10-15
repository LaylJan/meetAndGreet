import pandas as pd

# Read the Excel file
df = pd.read_excel("Attendees.xlsx")

# Rename columns for easier access
df.columns = ["Name", "Present", "Matched", "Lokal", "Gender", "Sports", "E-Games", "Reading", "Music", "Cooking", "Travel"]

# Filter out those who are not present or already matched
df = df[~((df["Present"].str.lower() == "no") | (df["Matched"].str.lower() == "yes"))]

# Split into males and females
males = df[df["Gender"].str.lower() == "male"].copy()
females = df[df["Gender"].str.lower() == "female"].copy()

# Keep track of who has already been matched
matched_males = set()
matched_females = set()

# Compatibility function
def compatibility_score(p1, p2):
    scores1 = p1[["Sports", "E-Games", "Reading", "Music", "Cooking", "Travel"]].astype(float).values
    scores2 = p2[["Sports", "E-Games", "Reading", "Music", "Cooking", "Travel"]].astype(float).values
    diff = abs(scores1 - scores2)
    score = 100 - (diff.mean() * 20)
    return round(score, 2)

print("💞 MATCH RESULTS 💞\n")

# Match each male to the most compatible female (only once)
for m_idx, male in males.iterrows():
    best_match = None
    best_score = -1

    for f_idx, female in females.iterrows():
        if female["Name"] in matched_females:
            continue  # skip already matched females

        score = compatibility_score(male, female)
        if score > best_score:
            best_score = score
            best_match = (female["Name"], female["Lokal"], score)

    if best_match:
        print(f"{male['Name']} ({male['Lokal']}) \t ❤️ \t {best_match[0]} ({best_match[1]}) → Compatibility: {best_match[2]}%")
        matched_males.add(male["Name"])
        matched_females.add(best_match[0])

print("\n✅ Matching complete! Everyone only matched once.")
