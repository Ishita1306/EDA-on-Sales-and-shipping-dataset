import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r'C:\Users\ishit\OneDrive\Python Dataset.xls'
df = pd.read_excel(file_path)

# Clean column names (remove extra spaces)
df.columns = df.columns.str.strip()

# Show columns to verify
print("Available columns:", df.columns.tolist())

# Convert dates
df['Order Date'] = pd.to_datetime(df['Order Date'], dayfirst=True)
df['Ship Date'] = pd.to_datetime(df['Ship Date'], dayfirst=True)

# Create new column for monthly analysis
df['Order Month'] = df['Order Date'].dt.to_period('M').astype(str)


# 1. Sales and Profit Trends Over Time
monthly = df.groupby('Order Month')[['Sales', 'Profit']].sum().reset_index()

plt.figure(figsize=(12, 6))
sns.lineplot(data=monthly, x='Order Month', y='Sales', label='Sales', color='dodgerblue', linewidth=2)
sns.lineplot(data=monthly, x='Order Month', y='Profit', label='Profit', color='orange', linewidth=2)
plt.fill_between(monthly['Order Month'], monthly['Sales'], monthly['Profit'], color='lightgray', alpha=0.5)
plt.title('Sales and Profit Trends Over Time', fontsize=16)
plt.xlabel('Month', fontsize=12)
plt.ylabel('Amount', fontsize=12)
plt.xticks(rotation=45)
plt.legend(title='Metrics', loc='upper left')
plt.tight_layout()
plt.show()


# 2. Top 10 States by Sales and Profit

top_states = df.groupby('State')[['Sales', 'Profit']].sum().sort_values('Sales', ascending=False).head(10)

plt.figure(figsize=(10, 6))
top_states['Sales'].plot(kind='barh', color='mediumseagreen', edgecolor='black', width=0.7)
top_states['Profit'].plot(kind='barh', color='coral', edgecolor='black', width=0.7, alpha=0.6)
plt.title('Top 10 States by Sales and Profit', fontsize=16)
plt.xlabel('Amount', fontsize=12)
plt.ylabel('State', fontsize=12)
plt.tight_layout()
plt.show()


# 3. Top 10 Cities by Sales
top_cities = df.groupby('City')['Sales'].sum().sort_values(ascending=False).head(10)

plt.figure(figsize=(10, 6))
sns.barplot(x=top_cities.index, y=top_cities.values, hue=top_cities.index, palette='viridis')

plt.title('Top 10 Cities by Sales', fontsize=16)
plt.xlabel('City', fontsize=12)
plt.ylabel('Sales', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


# 4. Segment-wise Sales Performance
plt.figure(figsize=(10, 6))
sns.violinplot(data=df, x='Segment', y='Sales', hue='Segment', palette='muted')

plt.title('Segment-wise Sales Performance', fontsize=16)
plt.xlabel('Customer Segment', fontsize=12)
plt.ylabel('Sales', fontsize=12)
plt.tight_layout()
plt.show()



# 5. Profit Distribution Across Discount Levels (Box Plot)
plt.figure(figsize=(10, 6))

sns.set(style="whitegrid")
sns.boxplot(data=df, x='Discount', y='Profit', hue='Discount', palette='Set2', width=0.6, fliersize=8, linewidth=2)

sns.scatterplot(data=df, x='Discount', y='Profit', color='red', s=80, alpha=0.6, marker='o')

plt.title('Profit Distribution Across Discount Levels (With Outliers)', fontsize=16)
plt.xlabel('Discount Percentage', fontsize=12)
plt.ylabel('Profit', fontsize=12)
plt.tight_layout()
plt.show()









