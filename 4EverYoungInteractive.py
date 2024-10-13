# Import necessary libraries
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from wordcloud import WordCloud, STOPWORDS
import difflib

# Function to display a word cloud with customizable stopwords
def display_wordcloud(text, title, additional_stopwords=None, max_words=100, colormap='viridis', background_color='white'):
    stopwords = set(STOPWORDS)
    if additional_stopwords:
        stopwords.update(additional_stopwords)
    
    wordcloud = WordCloud(
        width=800, height=400, background_color=background_color, colormap=colormap,
        stopwords=stopwords, max_words=max_words, min_font_size=10
    ).generate(text)
    
    fig, ax = plt.subplots()
    ax.imshow(wordcloud, interpolation='bilinear')
    ax.axis("off")
    ax.set_title(title, fontsize=15, fontweight='bold')
    st.pyplot(fig)

# Function to aggregate similar skills
def aggregate_skills(skill_list, reference_skills):
    aggregated_skills = []
    for skill in skill_list:
        closest_match = difflib.get_close_matches(skill, reference_skills, n=1)
        if closest_match:
            aggregated_skills.append(closest_match[0])
        else:
            aggregated_skills.append(skill)
    return aggregated_skills

# Load the cleaned student data
file_path = r"D:\Christ\Trimester 4\Web Project\student.xlsx"
data = pd.read_excel(file_path)

# Load the skill-to-question mapping
questions_file_path = r"D:\Christ\Trimester 4\Web Project\skill_questions.xlsx"
skill_questions = pd.read_excel(questions_file_path)

# Rename columns for easier access
data.rename(columns={
    'Research area/domain or  Projects that you have worked on or are currently working on (can be more than one project).\n\nIf there are more than one paper or project, kindly list it in points.': 'Research area/domain',
    'Your Achievements (can include valid certifications, leadership roles, sports achievements, volunteer work in any intra or inter-collegiate events, etc.)\nIf you have more than one achievement, please list them in points.': 'Achievements'
}, inplace=True)

# Streamlit App for batch and student analysis
st.title("Student Interview Dashboard")

# Sidebar for customizing visualizations
st.sidebar.title("Customization Options")
wordcloud_colormap = st.sidebar.selectbox("Select WordCloud Colormap", ['viridis', 'plasma', 'inferno', 'magma', 'cividis', 'Blues', 'Reds'])
max_words = st.sidebar.slider("Max Words in WordCloud", min_value=50, max_value=300, value=100)
background_color = st.sidebar.color_picker("WordCloud Background Color", value='#ffffff')
additional_stopwords = st.sidebar.text_area("Enter additional stopwords (comma-separated)")

# Split stopwords input into a list
if additional_stopwords:
    additional_stopwords = [word.strip().lower() for word in additional_stopwords.split(',')]
else:
    additional_stopwords = None

# Function to analyze all students in a selected batch
def batch_analysis(batch_year):
    batch_data = data[data['Batch Year'] == batch_year]
    
    st.write(f"## Analysis for Batch {batch_year}")
    
    # Skills Distribution for all students in the batch
    st.write("### Skills Distribution for the Batch")
    all_skills = batch_data['Skills you possess (Ex: Python, Power BI, R, SQL)'].str.cat(sep=',').lower()
    skills_list = [skill.strip() for skill in all_skills.split(',')]
    
    # Define common skills reference list to aggregate similar skills
    reference_skills = ['python', 'machine learning', 'data analysis', 'excel', 'sql', 'aws', 'html/css', 'power bi', 'r', 'tableau', 'java', 'c++']
    aggregated_skills_list = aggregate_skills(skills_list, reference_skills)
    
    # Count frequency of each skill
    skill_count = pd.Series(aggregated_skills_list).value_counts()
    
    # Filter skills that appear more than once to avoid clutter
    skill_count = skill_count[skill_count > 1]
    
    # Plot the cleaned skill count using an improved layout
    fig, ax = plt.subplots(figsize=(10, 8))  # Adjust the figure size
    sns.barplot(x=skill_count.values, y=skill_count.index, ax=ax, palette="Blues_d")
    ax.set_xlabel("Frequency", fontsize=12)
    ax.set_ylabel("Skills", fontsize=12)
    ax.set_title(f"Skill Popularity in Batch {batch_year}", fontsize=14, fontweight='bold')
    ax.tick_params(axis='y', labelsize=10)  # Increase label size
    st.pyplot(fig)
    
    # Display Word Cloud for Skills in the batch
    st.write(f"### Word Cloud for Batch {batch_year} Skills")
    all_skills_str = " ".join(aggregated_skills_list)
    display_wordcloud(all_skills_str, f"Skills Word Cloud for Batch {batch_year}", additional_stopwords, max_words, wordcloud_colormap, background_color)
    
    # Projects and Research Areas Word Cloud
    st.write(f"### Projects and Research Areas in Batch {batch_year}")
    all_projects = batch_data['Research area/domain'].str.cat(sep=' ').lower()
    display_wordcloud(all_projects, f"Projects and Research Word Cloud for Batch {batch_year}", additional_stopwords, max_words, wordcloud_colormap, background_color)

# Dropdown menu to select a batch
batch_year = st.selectbox("Select a Batch Year", data['Batch Year'].unique(), key='batch_selectbox')

# Perform analysis for the selected batch
if batch_year:
    batch_analysis(batch_year)

# Function to analyze a selected student
def student_analysis(student_id):
    student_data = data[data['Full Name'] == student_id].iloc[0]
    
    st.write(f"## Analysis for {student_data['Full Name']}")
    
    # Layout: split information in two columns
    col1, col2 = st.columns(2)
    
    with col1:
        st.write(f"**Email**: {student_data['Mail ID (University mail ID)']}")
        st.write(f"**Contact Number**: {student_data['Contact Number']}")
        st.write(f"**UG Degree**: {student_data['UG Degree']}")
        st.write(f"**Skills**: {student_data['Skills you possess (Ex: Python, Power BI, R, SQL)']}")
        st.write(f"**Projects and Research**: {student_data['Research area/domain']}")
        st.write(f"**Achievements**: {student_data['Achievements']}")
        st.write(f"**LinkedIn**: {student_data['Link/URL of your LinkedIn account']}")
        st.write(f"**GitHub**: {student_data['Link/URL of your GitHub']}")

    with col2:
        # Example Visualization: Skills Count
        skills = student_data['Skills you possess (Ex: Python, Power BI, R, SQL)'].split(',')
        skills = [skill.strip().lower() for skill in skills]
        skill_count = pd.Series(skills).value_counts()
        
        st.write("### Skills Distribution")
        fig, ax = plt.subplots()
        sns.barplot(x=skill_count.values, y=skill_count.index, ax=ax, palette="Blues_d")
        ax.set_xlabel("Frequency", fontsize=12)
        ax.set_ylabel("Skills", fontsize=12)
        ax.set_title("Skill Popularity", fontsize=14, fontweight='bold')
        st.pyplot(fig)
        
        # Display Word Cloud for Skills
        st.write("### Skills Word Cloud")
        all_skills = " ".join(skills)
        display_wordcloud(all_skills, "Skills Word Cloud", additional_stopwords, max_words, wordcloud_colormap, background_color)

    # Display Interview Questions Related to Skills
    st.write("### Relevant Interview Questions")
    for skill in skills:
        related_questions = skill_questions[skill_questions['Skill'].str.lower() == skill]['Interview Questions']
        if not related_questions.empty:
            st.write(f"**For skill '{skill}'**:")
            for question in related_questions:
                st.write(f"- {question}")
        else:
            st.write(f"No questions found for skill '{skill}'")

# Dropdown menu to select a student
student_id = st.selectbox("Select a student", data['Full Name'].unique(), key='student_selectbox')

# Perform analysis for the selected student
if student_id:
    student_analysis(student_id)
