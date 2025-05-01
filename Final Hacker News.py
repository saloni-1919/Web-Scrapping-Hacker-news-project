"""
Introduction to coding IS 612
Date: 30th April, 2025
Purpose: Final Project: Hacker News using Web Scraping and Data Analysis using beautiful soup

  Beautiful Soup is a library that makes it easy to scrape information from web pages. 
  It sits atop an HTML or XML parser, providing Pythonic idioms for iterating, searching, and modifying the parse tree.

Team Members:
- Saloni Sanjaybhai Nathani
- Ayanda Timothy Kutobwa
- Desai, Ishika Vijaykumar
- Vempalli, Prashanth
"""
import requests                # Importing requests library to send HTTP requests and fetch web content
from bs4 import BeautifulSoup    # Importing BeautifulSoup to parse and navigate HTML data from web pages
import pandas as pd    # Importing pandas for data manipulation and analysis using DataFrames
import matplotlib.pyplot as plt    # Importing matplotlib for creating plots and charts
import seaborn as sns     # Importing seaborn for advanced data visualization and statistical plotting
import os                # Importing os module to interact with the operating system (e.g., checking if a file exists)

#Please read the text file 'readme.txt' for technical requirements of this project. 


#The requirements this project has achieved:

# It must pull data from the internet (typically webpages) ---> Marked as point 1 in the code below 

# It must pull data from 5 different webpages (using pandas) --> Marked as point 2 in the code below

# It must do some basic calculations on numerical data from 3 of those pages (mean, median, mode, min, max) --> Marked as point 3 in the code below

# It must make use of your own methods where appropriate --> Marked as point 4 in the code below

# It must produce at least 3 charts from 3 of those pages --> Marked as point 5 in the code below

# It must save the data to a file --> Marked as point 6 in the code below

# Additional requirements achieved (must have 4 of the following, more than 4 is extra credit)

# 1 of the webpages includes data which changes daily  --> Marked as point 7 in the code below

# The 5 webpages are from 5 different websites  --> Marked as point 8 in the code below

# It does basic calculations on all 5 of those webpages  --> Marked as point 9 in the code below

# It uses BeautifulSoup on 1 of the webpages and parses the html --> Marked as point 10 in the code below

# It produces 5 charts from all 5 of the webpages --> Marked as point 11 in the code below

# It saves the charts in addition to the data, to image files  --> Marked as point 12 in the code below

# It reads the previous file, and notes if any webpage has changed data* (hard) --> Marked as point 13 in the code below

# ----------------- SCRAPING HACKER NEWS -----------------

def scrape_hackernews():
    """Scrapes top stories from Hacker News with their scores."""
    # Following lines are necessary to access the website to Scrape without getting blocked.
    
    #point  1 
    url = "https://news.ycombinator.com/"           # Defining the URL of Hacker News homepage
    headers = {"User-Agent": "Mozilla/5.0"}         # Setting headers to mimic a real browser request
    response = requests.get(url, headers=headers)   # Sending HTTP GET request to the URL
    
    if response.status_code != 200:     # Checking if the request was successful   200 is the status code for successful loading of website
        print(f"Failed to retrieve Hacker News. Status code: {response.status_code}")
        return None
    
    # Point 10  uses BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')  # Parsing the HTML content using BeautifulSoup
    stories = []                                           # Defining a list to store story data
    rows = soup.find_all('tr', class_='athing')             # Getting all story rows with class "athing"
    
    for row in rows:      # Iterating through each story row
        title_elem = row.find('span', class_='titleline')  # Getting the title element
        score_elem = row.find_next_sibling('tr').find('span', class_='score')  # Getting the score element from the next sibling row
        
        if title_elem:        # Checking if title exists
            title = title_elem.text.strip()     # Extracting title text
            link = title_elem.find('a')['href']     # Getting the link URL
            score = int(score_elem.text.split()[0]) if score_elem else 0  # Default score 0   # Getting the score or assigning 0 if not found
            
            stories.append([title, link, score])   # Appending the story data to the list
            
    # Point 2
    df = pd.DataFrame(stories, columns=['Title', 'Link', 'Score'])  # Creating a DataFrame from the list of stories
    return df

# ----------------- BASIC CALCULATIONS -----------------

# Point 3  basic calculations on numerical data from 3 of those pages
# Point 9 basic calculations on all 5 of those webpages
def save_scores_and_stats_to_excel(df):     # (   pip install XlsxWriter  ) (required for xlsx file which is for saving basic calulations)
    """Saves story titles with scores and statistical summary into one Excel sheet and returns the stats DataFrame."""
   
    df_sorted = df.sort_values(by='Score', ascending=False)      # Sorting the DataFrame by score in descending order

    df_sorted = df_sorted[['Title', 'Score']]     # Selecting only Title and Score columns

    # Calculating statistics
    total_score = df_sorted['Score'].sum()      # Calculating total score
    mean_score = round(df_sorted['Score'].mean(), 2)    # Calculating mean score
    median_score = df_sorted['Score'].median()           # Calculating median score
    mode_score = df_sorted['Score'].mode()[0] if not df_sorted['Score'].mode().empty else None   # Calculating mode score
    min_score = df_sorted['Score'].min()     # Getting minimum score
    max_score = df_sorted['Score'].max()      # Getting maximum score 
   
         
    stats_data = {     # Creating a dictionary of statistics
        'Statistic': ['Total Score', 'Mean Score', 'Median Score', 'Mode Score', 'Min Score', 'Max Score'],
        'Value': [total_score, mean_score, median_score, mode_score, min_score, max_score]
    }
    stats_df = pd.DataFrame(stats_data)    # Creating a DataFrame from statistics

    # Creating Excel writer and writing both DataFrames
    #point 6 save data to file
    with pd.ExcelWriter('hackernews_full_summary.xlsx', engine='xlsxwriter') as writer:
        df_sorted.to_excel(writer, index=False, sheet_name='Summary', startrow=0)    # Writing the sorted stories to Excel
        stats_df.to_excel(writer, index=False, sheet_name='Summary', startrow=len(df_sorted) + 3)    # Writing the statistics below the stories

    print("✅ Full data and stats saved to 'hackernews_full_summary.xlsx'")  #printing saved file message

    # Return the stats DataFrame
    return stats_data  

# ----------------- PLOTTING FUNCTIONS -----------------
# point 5 produce at least 3 charts from 3 of those pages
# Point 11 produces 5 charts from all 5 of the webpages
def plot_hackernews(df):     # (   pip install matplotlib   ) (   pip install seaborn   )
    """Plots 3 different charts for Hacker News data."""
    # Get top 10 highest-scoring stories
    top_10 = df.nlargest(10, 'Score')

    # --- Bar Chart ---
    top_10['Title'] = top_10['Title'].apply(lambda x: ' '.join(x.split()[:3]))

    plt.figure(figsize=(12, 6))     # Creating bar chart
    sns.barplot(x=top_10['Title'], y=top_10['Score'], palette='coolwarm')     #assigning x and y axis
    plt.ylabel("Upvotes (Score)")  # Y-axis label
    plt.xlabel("Story Title")  # X-axis label
    plt.title("Top 10 Hacker News Stories by Score")
    plt.xticks(rotation=45, ha='right')  # Rotating x-axis labels for better readability
    plt.grid(axis='y', linestyle='--', alpha=0.7)  # Grid for y-axis
    plt.savefig('top_10_hackernews_bar_chart.png')  #point 6 and 12 save chart to file
    plt.show()

    # --- Histogram of Scores ---
    plt.figure(figsize=(10, 5))   #figure size
    sns.histplot(df['Score'], bins=20, kde=True, color='blue')   # Plotting histogram with 20 bins and kernel density estimate
    plt.xlabel("Upvotes (Score)")    # Setting label for X-axis
    plt.ylabel("Number of Stories")    # Setting label for Y-axis
    plt.title("Distribution of Hacker News Story Scores")    # Setting chart title
    plt.grid(axis='y', linestyle='--', alpha=0.7)     # Displaying horizontal grid lines for readability
    plt.savefig('hackernews_score_distribution.png') #Point 6 and 12 save chart to file
    plt.show()   # Displaying the chart 
      
    # --- Pie Chart for Top 10 Scores ---
    plt.figure(figsize=(8, 8))    #figure size
    plt.pie(top_10['Score'], labels=top_10['Title'], autopct='%1.1f%%', colors=sns.color_palette('pastel'))   # Plotting pie chart
    plt.title("Top 10 Stories Share of Total Score")  # Setting chart title
    plt.savefig('top_10_pie_chart.png')    #Point 6 and 12 save chart to file
    plt.show()   # Displaying the chart 
    
    # --- Line Chart for Scores ---
    df_sorted = df.sort_values(by='Score', ascending=False).reset_index(drop=True)   # Sorting stories by score in descending order and resetting index
    plt.figure(figsize=(10, 5))    # Creating a new figure with width 10 and height 5
    plt.plot(df_sorted.index + 1, df_sorted['Score'], marker='o', color='green')    # Plotting line chart with green dots for each story's score
    plt.title("Line Chart of Hacker News Story Scores")   # Setting the chart title
    plt.xlabel("Story Rank (by Score)")   # Labeling the X-axis
    plt.ylabel("Score")       # Labeling the Y-axis
    plt.grid(True)         # Enabling grid lines for better readability
    plt.savefig("score_line_chart.png")  #Point 6 and 12 save chart to file
    plt.show()   # Displaying the chart 

    # --- Box Chart for Scores ---
    plt.figure(figsize=(8, 4))   #figure size
    sns.boxplot(x=df['Score'], color='orange')   #plotting boc plot chart
    plt.title("Box Plot of Hacker News Story Scores")   # Setting the chart title
    plt.xlabel("Score")   # Labeling the X-axis
    plt.savefig("score_boxplot.png")    #Point 6 and 12 save chart to file
    plt.show()   # Displaying the chart 


# ----------------- SAVING DATA AND FILE CHECK -----------------

#Point 6 save data to file
#point 4 appropriate methods
#point 7 webpage data that changes daily
#point 8  5 different webpages 

# Point 13 It reads the previous file, and notes if any webpage has changed data
def save_data(df):#function for updating and getting insights of files and live data changes.
    """Saves scraped data to a CSV file and checks for additions, deletions, score changes, and reshuffling."""
    file_path = 'hackernews_data.csv'    # Defining the file path where data will be saved

    if os.path.exists(file_path):    # Checking if the previous data file exists
        old_df = pd.read_csv(file_path)    # Reading old data from CSV file

        old_titles = set(old_df['Title'])     # Extracting old story titles into a set
        new_titles = set(df['Title'])         # Extracting new story titles into a set

        # Detect added and removed stories
        added_titles = new_titles - old_titles    # Finding new titles not present in old data
        removed_titles = old_titles - new_titles    # Finding titles that were removed

        # Newly Added stories
        if added_titles:
            print("🆕 Added stories:")   # Print header for added stories
            for title in added_titles:   
                score = df[df['Title'] == title]['Score'].values[0]     # Getting score of the added story
                print(f"   → {title[:60]}... (Score: {score})")         # Printing title and score

        # Removed stories
        if removed_titles:
            print("\n❌ Removed stories:")     # Print header for removed stories
            for title in removed_titles:
                score = old_df[old_df['Title'] == title]['Score'].values[0]    # Getting score of the removed story
                print(f"   → {title[:60]}... (Score: {score})")      # Printing title and score

        # Score updates
        common_titles = new_titles & old_titles    # Finding stories that are common in both old and new data
        score_changes = []       # List to hold titles with updated scores
        reshuffled = []       # List to hold titles with same score but possible reordering


        for title in common_titles:   #iterating over titles
            old_score = old_df[old_df['Title'] == title]['Score'].values[0]    # Old score of the story
            new_score = df[df['Title'] == title]['Score'].values[0]            # New score of the story
            if old_score != new_score:  
                score_changes.append((title, old_score, new_score))    # Add to score_changes if score has changed
            else:
                reshuffled.append(title)    # Add to reshuffled if score unchanged

        if score_changes:
            print("\n🔁 Score updates:")  # Print header for updated scores
            for title, old, new in score_changes:
                print(f"   → {title[:60]}... (Score: {old} → {new})")      # Print title and updated score info

        if reshuffled:
            print("\n🔀 Stories reshuffled (no score change):")   # Print header for reshuffled stories
            for title in reshuffled[:5]:  # limiting to top 5 reshuffles
                print(f"   → {title[:60]}")

    else:
        print("📥 Saving new data to file (no previous file found).")   # Informing user if no old data was found

    # Saveing the latest data
    df.to_csv(file_path, index=False)


# ----------------- MAIN FUNCTION -----------------

def main():
    # Scraping Hacker News
    df = scrape_hackernews()
    
    if df is not None and not df.empty:
        # Calculating statistics and saving to Excel
        stats = save_scores_and_stats_to_excel(df)
        
        # Printing statistics
        print("\nHacker News Statistics:")
        for key, value in stats.items():
            print(f"{key}: {value}")   #printing key values
        
        # Ploting the data
        plot_hackernews(df)
        
        # Saving data and check for changes
        save_data(df)
    else:
        print("Failed to scrape data or data is empty.")

# Running the main function
if __name__ == "__main__":
    main()
