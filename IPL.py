import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl

#Load the file
filepath="ipl_players.xlsx"
df=pd.read_excel(filepath)
"Load excel file "

#Data Cleaning:
"cleaning data by using the inbuilt functions isnull().sum(),drop_duplicates(),fillna(0),astype(datatype)"

#1.find the missing values:
def missing_values(df):
    print("Missing values column wise:\n")
    missing = df.isnull().sum()
    print(missing)
    return missing

#2.Drop Duplicate rows
def drop_Duplicate_rows(df):
    print("Dropping duplicate rows:\n")
    df_cleaned = df.drop_duplicates()
    return df_cleaned

#3.Change datatatype:
def changetype(df):
  print("Change the dtatype of Matches column...\n")
  df['Matches'] = df['Matches'].fillna(0).astype(int)
  return df

#4.drop or fill null values
def simple_fillna(df):
    df = df.fillna(0)
       
    print("Fill NULL values\n")
    return df

#Data Analysis:
"Data analysis means to perform the operation on data using the userdefine or inbuilt functions like find average,mean etc."
def analyze_data(df):
    #1.Average price of palyers:
    print("Average Price:", df['Price'].mean())
    
    #2.count the players based on brand:
    print("Brand Distribution:\n", df['Brand'].value_counts())
    
    #3.maximum rating of the player:
    print("Max Rating:", df['Rating'].max())
    
    #4.minimum rating of the player:
    print("Min Rating:", df['Rating'].min())
    
    #Group the values acoording to the team
    grouped = df.groupby('Team')['Runs'].sum()
    print("Total Runs by Team:\n", grouped)
    
    #retiurn the sort data to the main
    return df.sort_values(by='Rating', ascending=False)
    print("Sort the data According to the rating\n")

#save the modified file:
def save_data(df, output_path):
   df.to_excel(output_path, index=False)

# Charts:
#Bar chart:
def plot_brand_distribution(df):
    df['Brand'].value_counts().plot(kind='bar', title='Brand Distribution')
    "we count brand means no of brands"
    plt.xlabel('Brand')
    plt.ylabel('Count')
    plt.tight_layout()
    plt.show()

#Pie Chart:
def plot_rating_distribution(df):
    df['Rating'].value_counts().plot(kind='pie',autopct='%0.1f%%',title='Rating Distribution')
    plt.ylabel('')
    plt.tight_layout()
    plt.show()

#Line chart:
def plot_runs_vs_matches(df):
    plt.plot(df['Matches'], df['Runs'], marker='o')
    plt.title('Runs vs Matches')
    plt.xlabel('Matches')
    plt.ylabel('Runs')
    plt.grid(True)
    plt.show()

#Scatter plot:
def plot_matches_by_player(df):
    sns.scatterplot(x='Player', y='Matches', data=df)
    plt.title('Matches Played by Each Player')
    plt.xlabel('Player')
    plt.ylabel('Matches')
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.show()

#Main Function:
def main(df):

    missing_values(df)
    df=drop_Duplicate_rows(df)
    df=changetype(df)
    df=simple_fillna(df)
    df_sorted = analyze_data(df)
    save_data(df_sorted, 'ipl_players_cleaned.xlsx')
    #charts
    plot_brand_distribution(df)
    plot_rating_distribution(df)
    plot_runs_vs_matches(df)
    plot_matches_by_player(df)

if __name__ == "__main__":
    main(df)
