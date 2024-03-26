import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load your data
# df = pd.read_csv('data.csv')

def summarize_dataframe(df):
    print(f"Number of rows: {df.shape[0]}")
    print(f"Number of columns: {df.shape[1]}")
    print("\nColumn Names: ", df.columns.tolist())
    print("\nData Types:")
    print(df.dtypes)
    print("\nMissing Values:")
    print(df.isnull().sum())
    print("\nNumber of Unique Values:")
    print(df.nunique())
    print(f"\nNumber of duplicate rows: {df.duplicated().sum()}")

    # For categorical (object) columns
    categorical_columns = df.select_dtypes(include=['object']).columns
    for col in categorical_columns:
        print(f"\nColumn: {col}")
        print(f"Number of categories: {df[col].nunique()}")
        print(f"Most frequent category: {df[col].mode()[0]}")

    # For numerical columns
    numerical_columns = df.select_dtypes(include=['int64', 'float64']).columns
    for col in numerical_columns:
        print(f"\nColumn: {col}")
        print(f"Mean: {df[col].mean()}")
        print(f"Median: {df[col].median()}")
        print(f"Standard Deviation: {df[col].std()}")

    # For datetime columns
    datetime_columns = df.select_dtypes(include=['datetime64']).columns
    for col in datetime_columns:
        print(f"\nColumn: {col}")
        df[col] = pd.to_datetime(df[col])
        print(f"Earliest date: {df[col].min()}")
        print(f"Latest date: {df[col].max()}")
        print(f"Number of unique dates: {df[col].nunique()}")

# Call the function to summarize the data
# summarize_dataframe(df)

# Let's define a function to plot the distribution of categorical columns


def plot_sub_categorical(df, column, top_n=20, max_label_length=30):
    plt.figure(figsize=(12, 6))
    
    # Select top_n categories by count
    top_categories = df[column].value_counts().index[:top_n]
    df_top_n = df[df[column].isin(top_categories)].copy()

    # Truncate category names to a maximum length
    df_top_n[column] = df_top_n[column].apply(lambda x: (x[:max_label_length] + '...') if len(x) > max_label_length else x)

    # Create the count plot
    plot = sns.countplot(data=df_top_n, y=column, order=df_top_n[column].value_counts().index, color='skyblue')
    plt.title(f'Distribution of {column}')
    plt.xlabel('Count')
    plt.ylabel(column)
    
    # Add labels to the bars
    for p in plot.patches:
        width = p.get_width()
        plt.text(width + 0.3,  # Add a small buffer for the text
                 p.get_y() + p.get_height() / 2,
                 '{:1.0f}'.format(width), 
                 va='center') 

    plt.show()
    

def plot_categorical(df):
    
    # List of categorical columns to analyze
    categorical_columns = df.select_dtypes(include=['object']).columns

    # Analyze each categorical column
    for column in categorical_columns:
        plot_sub_categorical(df, column)

