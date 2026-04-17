import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# Load the dataset using openpyxl for the Excel file
file_path = "Clemson Football Tickets  (Responses).xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# Clean up column names by stripping any leading/trailing whitespace
df.columns = df.columns.str.strip()

# Set the visual style for the plots
sns.set_theme(style="whitegrid")
clemson_colors = ['#F56600', '#522D80'] # Orange and Regalia

# ==========================================
# Visual 1: Disparity in Student Attendance 
# (Supports Problem & "Guaranteed Baseline Access")
# ==========================================
plt.figure(figsize=(9, 6))
attendance_col = "How many Clemson Football games do you regularly attend in a single season?"

# Function to extract numbers from the free-text responses
def clean_attendance(val):
    val = str(val).lower()
    if 'all' in val or 'every' in val:
        return 7 # Assuming ~7 home games in a season
    nums = re.findall(r'\d+', val)
    if nums:
        return int(nums[0])
    return None

df['Cleaned_Attendance'] = df[attendance_col].apply(clean_attendance)

# Drop missing/unparsable values just for the plot
plot_data_v1 = df.dropna(subset=['Cleaned_Attendance'])

sns.countplot(data=plot_data_v1, x='Cleaned_Attendance', color='#F56600')
plt.title("Current Disparity in Student Game Attendance\n(Highlights the Need for Guaranteed Baseline Access)", fontsize=14, fontweight='bold')
plt.xlabel("Number of Games Attended in a Season", fontsize=12)
plt.ylabel("Number of Students", fontsize=12)

plt.savefig("attendance_disparity_bar.png", bbox_inches='tight')
print("Visual 1 saved as 'attendance_disparity_bar.png'")
plt.close()

# ==========================================
# Visual 2: Impact of Easier Tickets on Attendance (Pie Chart)
# (Supports Section 5.1 Call to Action)
# ==========================================
plt.figure(figsize=(8, 6))
easier_tickets_col = "If tickets were easier to get, would you attend more games?"

# Count the Yes/No responses
attendance_impact_counts = df[easier_tickets_col].astype(str).str.strip().str.lower().value_counts()

plt.pie(
    attendance_impact_counts, 
    labels=attendance_impact_counts.index.str.capitalize(), 
    autopct='%1.1f%%', 
    startangle=140, 
    colors=clemson_colors,
    textprops={'fontsize': 18}
)

plt.title("Would Students Attend More Games if Tickets Were Easier to Get?", fontsize=14, fontweight='bold')
plt.axis('equal') 

plt.savefig("easier_tickets_impact_pie.png", bbox_inches='tight')
print("Visual 2 saved as 'easier_tickets_impact_pie.png'")
plt.close()

# ==========================================
# Visual 3: Sentiment on Ticket Resellers (Donut Chart)
# (Supports "Market Regulation" Criteria)
# ==========================================
plt.figure(figsize=(8, 8)) # Back to a square since the legend is inside

# Hardcoded data based on manual analysis
labels = [
    'Negative: Overpricing (25)', 
    'Negative: Other (4)', 
    'Neutral / Positive (11)'
]
sizes = [25, 4, 11]

# Colors: Clemson Regalia (Purple) for the main negative, a muted purple for other negative, Orange for Neutral/Positive
colors = ['#522D80', '#8573A1', '#F56600']

# Create the pie chart base
wedges, texts, autotexts = plt.pie(
    sizes, 
    autopct='%1.1f%%', 
    startangle=140, 
    colors=colors,
    pctdistance=0.85, # Pushed out slightly to make room for the inner legend
    wedgeprops={'edgecolor': 'white', 'linewidth': 2},
    textprops={'fontsize': 20, 'fontweight': 'bold', 'color': 'white'}
)

# Draw a slightly larger white circle in the center to fit the legend
centre_circle = plt.Circle((0, 0), 0.65, fc='white')
fig = plt.gcf()
fig.gca().add_artist(centre_circle)

# Add the legend perfectly inside the donut hole
plt.legend(
    wedges, 
    labels, 
    title="Sentiments", 
    loc="center", # Centers the legend in the hole
    fontsize=18,
    title_fontsize=20,
    frameon=False # Removes the box around the legend for a clean look
)

plt.title("Student Sentiment Toward Ticket Resellers", fontsize=24, fontweight='bold')
plt.axis('equal')

plt.tight_layout() 
plt.savefig("reseller_sentiment_donut.png", bbox_inches='tight')
print("Visual 3 saved as 'reseller_sentiment_donut.png'")
plt.close()