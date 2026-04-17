import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset using openpyxl for the Excel file
file_path = "Clemson Football Tickets  (Responses).xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# Clean up column names by stripping any leading/trailing whitespace
df.columns = df.columns.str.strip()

# Set the visual style for the plots
sns.set_theme(style="whitegrid")
# Define Clemson colors
clemson_colors = ['#F56600', '#522D80'] # Orange and Regalia

# ==========================================
# Visual 1: Participation in the Ticket Lottery
# ==========================================
plt.figure(figsize=(8, 6))

# Define the exact column name from your Excel file
lottery_col = "Do you participate in the Clemson Football Ticket Lottery?"

# Clean and count the responses (ignoring case and whitespace)
participation_counts = df[lottery_col].astype(str).str.strip().str.lower().value_counts()

# Create a pie chart
plt.pie(
    participation_counts, 
    labels=participation_counts.index.str.capitalize(), 
    autopct='%1.1f%%', 
    startangle=140, 
    colors=clemson_colors 
)
plt.title("Participation in the Clemson Football Ticket Lottery", fontsize=14, fontweight='bold')
plt.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.

# Save the first visual
plt.savefig("lottery_participation_pie.png", bbox_inches='tight')
print("Visual 1 saved as 'lottery_participation_pie.png'")

# ==========================================
# Visual 2: Impact of Easier Tickets on Attendance
# ==========================================
plt.figure(figsize=(8, 6))

# Define the exact column name from your Excel file
easier_tickets_col = "If tickets were easier to get, would you attend more games?"

# Clean and count the responses (ignoring case and whitespace)
attendance_impact_counts = df[easier_tickets_col].astype(str).str.strip().str.lower().value_counts()

# Create a pie chart
plt.pie(
    attendance_impact_counts, 
    labels=attendance_impact_counts.index.str.capitalize(), 
    autopct='%1.1f%%', 
    startangle=140, 
    colors=clemson_colors
)

plt.title("Would Students Attend More Games if Tickets Were Easier to Get?", fontsize=14, fontweight='bold')
plt.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.

# Save the second visual
plt.savefig("easier_tickets_impact_pie.png", bbox_inches='tight')
print("Visual 2 saved as 'easier_tickets_impact_pie.png'")