import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl

filepath="Sales Forecasting.xlsx"
df=pd.read_excel(filepath)
print(df)

#Data Cleaning
def missing_values(df):
    print("Missing values column wise\n")
    missing= df.isnull().sum()
    print(missing)
    return missing

def drop_duplicate(df):
    print("Droping duplicate rows:\n")
    df_cleaned=df.drop_duplicates()
    return df_cleaned

def changetype(df):
    print("change datatype:\n")
    df=df.copy()
    df['Unit_Sold']=df['Unit_Sold'].fillna(0).astype(float)
    return df

def fillnull(df):
    df=df.fillna(0)
    print("fill Null values:\n")
    return df

#Data Analysis
def analyze_data(df):
    print("Average Selling price",df['Selling_Price'].mean())

    print("Count the no of product",df['Product_Name'].value_counts())

    print("maximum selling price",df['Total_Revenue'].max())

    print("Minimum Selling price",df['Unit_Sold'].min())

    grouped=df.groupby('Category')['Product_Name'].count()
    print("total product by category:\n",grouped)

    return df.sort_values(by='Total_Revenue',ascending=False)

def save_data(output_path):
    df.to_excel(output_path,index=False)

data = df.groupby('State')['Product_Name'].count()
data.plot(kind='bar', color='Blue')
plt.title("Product Sales by State")
plt.xlabel("States")
plt.ylabel("Number of Products Sold")
plt.tight_layout()
plt.show()

#Bar Chart
def product_total_Revenue(df):
    revenue_by_product = df.groupby('Category')['Total_Revenue'].sum()
    revenue_by_product.plot(kind='pie',autopct='%0.1f%%',title='Product Share by Total Revenue',startangle=90)
    plt.ylabel('')  
    plt.tight_layout()
    plt.show()


#Pie Chart
def main(df):
    missing_values(df)
    df=drop_duplicate(df)
    df=changetype(df)
    df=fillnull(df)
    df=analyze_data(df)
    product_total_Revenue(df)
    save_data('Sales_Forecasting_cleaned.xlsx')

if __name__=="__main__":
    main(df)

