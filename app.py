import streamlit as st
from groq import Groq
from supabase import create_client, Client
from docx import Document
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import io

# ─────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────
GROQ_API_KEY    = st.secrets["GROQ_API_KEY"]
SUPABASE_URL    = st.secrets["SUPABASE_URL"]
SUPABASE_KEY    = st.secrets["SUPABASE_KEY"]

groq_client = Groq(api_key=GROQ_API_KEY)
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# ─────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(page_title="AI Teacher Assistant", page_icon="📚", layout="wide")

# Make the right "Actions" column stick to the viewport while scrolling
st.markdown("""
<style>
    div[data-testid="column"]:nth-of-type(2) {
        position: sticky;
        top: 4rem;
        align-self: flex-start;
        max-height: calc(100vh - 5rem);
        overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────
for key, val in {
    "user": None,
    "token": None,
    "refresh_token": None,
    "current_output": "",
    "current_title": "",
    "current_task": "",
    "edit_mode": False,
    "edit_id": None,
    "saved_plans": [],
    "auth_mode": "login",
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# Re-attach session on every script rerun (Streamlit creates a fresh
# supabase client each run, so RLS won't know who the user is otherwise)
if st.session_state.user and st.session_state.token and st.session_state.refresh_token:
    try:
        supabase.auth.set_session(st.session_state.token, st.session_state.refresh_token)
    except Exception:
        pass

# ─────────────────────────────────────────
#  GRADE OPTIONS
# ─────────────────────────────────────────
GRADE_OPTIONS = {
    "IB (International Baccalaureate)":  ["Grade 11", "Grade 12"],
    "IGCSE (Cambridge)":                  ["Grade 9", "Grade 10"],
    "A-Level (Cambridge)":                ["Grade 11", "Grade 12"],
    "Cambridge Checkpoint":               ["Grade 5", "Grade 6", "Grade 7", "Grade 8"],
    "Singapore (MOE)":                    ["K1", "K2"] + [f"Grade {i}" for i in range(1, 9)],
    "CBSE (India)":                       [f"Grade {i}" for i in range(1, 13)],
    "American (Common Core)":            [f"Grade {i}" for i in range(1, 13)],
    "Australian (ACARA)":                 [f"Grade {i}" for i in range(1, 13)],
    "British (National Curriculum)":      [f"Year {i}" for i in range(1, 14)],
}

# ─────────────────────────────────────────
#  SUBJECT OPTIONS (per curriculum)
# ─────────────────────────────────────────
SUBJECT_OPTIONS = {
    "IB (International Baccalaureate)":  ["Computer Science", "Mathematics", "Physics", "Chemistry",
                                           "Biology", "English", "Economics", "Business Management",
                                           "Psychology", "History", "Geography"],
    "IGCSE (Cambridge)":                  ["Computer Science", "Mathematics", "Physics", "Chemistry",
                                           "Biology", "English", "Economics", "Business Studies",
                                           "ICT", "History", "Geography"],
    "A-Level (Cambridge)":                ["Computer Science", "Mathematics", "Physics", "Chemistry",
                                           "Biology", "English", "Economics", "Business",
                                           "Psychology", "History", "Geography"],
    "Cambridge Checkpoint":               ["Mathematics", "Science", "English", "ICT", "Computer Science"],
    "Singapore (MOE)":                    ["Mathematics", "Science", "English", "Computer Science",
                                           "Social Studies", "Art"],
    "CBSE (India)":                       ["Mathematics", "Science", "English", "Computer Science",
                                           "Social Science", "Hindi", "Physics", "Chemistry", "Biology"],
    "American (Common Core)":            ["Mathematics", "Science", "English Language Arts",
                                           "Computer Science", "Social Studies"],
    "Australian (ACARA)":                 ["Mathematics", "Science", "English", "Digital Technologies",
                                           "Humanities and Social Sciences"],
    "British (National Curriculum)":      ["Mathematics", "Science", "English", "Computing",
                                           "History", "Geography"],
}

# ─────────────────────────────────────────
#  TOPIC OPTIONS (per subject)
# ─────────────────────────────────────────
TOPIC_OPTIONS = {
    "Computer Science": [
        "Abstraction & Decomposition", "Algorithms & Flowcharts", "Programming Basics (Variables, Data Types)",
        "Sequencing, Selection & Iteration", "Functions & Procedures", "Arrays & Lists",
        "Searching & Sorting Algorithms", "Data Structures (Stacks, Queues, Trees)",
        "Object-Oriented Programming", "Databases & SQL", "Networks & Protocols",
        "Binary, Hex & Number Systems", "Logic Gates & Boolean Algebra", "Cyber Security",
        "Web Development (HTML/CSS/JS)", "Python Programming", "Computational Thinking",
        "Software Development Life Cycle", "Ethics in Computing", "Artificial Intelligence Basics"
    ],
    "Mathematics": [
        "Number & Place Value", "Fractions, Decimals & Percentages", "Algebra & Expressions",
        "Equations & Inequalities", "Geometry & Shapes", "Measurement & Units",
        "Statistics & Data Handling", "Probability", "Ratio & Proportion",
        "Functions & Graphs", "Trigonometry", "Calculus (Differentiation/Integration)",
        "Vectors & Transformations", "Sequences & Series", "Coordinate Geometry"
    ],
    "Physics": [
        "Motion & Forces", "Energy & Work", "Electricity & Circuits", "Waves & Sound",
        "Light & Optics", "Magnetism & Electromagnetism", "Thermal Physics",
        "Nuclear Physics & Radioactivity", "Gravitation", "Pressure & Density"
    ],
    "Chemistry": [
        "Atomic Structure & Periodic Table", "Bonding (Ionic, Covalent, Metallic)",
        "Chemical Reactions & Equations", "Acids, Bases & Salts", "Rates of Reaction",
        "Energetics", "Organic Chemistry", "Electrolysis", "States of Matter", "Stoichiometry"
    ],
    "Biology": [
        "Cell Structure & Function", "Nutrition & Digestion", "Respiration", "Photosynthesis",
        "Transport in Plants & Animals", "Reproduction", "Genetics & Inheritance",
        "Evolution & Natural Selection", "Ecosystems & Ecology", "Human Body Systems"
    ],
    "English": [
        "Reading Comprehension", "Creative Writing", "Grammar & Punctuation",
        "Poetry Analysis", "Persuasive Writing", "Narrative Writing", "Shakespeare",
        "Novel Study", "Speaking & Listening", "Vocabulary Building"
    ],
    "English Language Arts": [
        "Reading Comprehension", "Creative Writing", "Grammar & Punctuation",
        "Poetry Analysis", "Persuasive Writing", "Narrative Writing", "Novel Study",
        "Speaking & Listening", "Vocabulary Building"
    ],
    "Science": [
        "Living Things & Habitats", "Forces & Motion", "Materials & Their Properties",
        "Earth & Space", "Energy", "States of Matter", "Human Body", "Plants",
        "Electricity", "Sound & Light"
    ],
    "Economics": [
        "Demand & Supply", "Market Structures", "Macroeconomics & GDP", "Inflation & Unemployment",
        "International Trade", "Government Intervention", "Microeconomics Basics",
        "Elasticity", "Money & Banking", "Economic Development"
    ],
    "Business Management": [
        "Business Organisation", "Marketing", "Finance & Accounts", "Human Resource Management",
        "Operations Management", "Business Strategy", "Entrepreneurship", "External Environment"
    ],
    "Business Studies": [
        "Business Organisation", "Marketing", "Finance & Accounts", "Human Resource Management",
        "Operations Management", "Business Strategy", "Entrepreneurship", "External Environment"
    ],
    "Business": [
        "Business Organisation", "Marketing", "Finance & Accounts", "Human Resource Management",
        "Operations Management", "Business Strategy", "Entrepreneurship", "External Environment"
    ],
    "Psychology": [
        "Research Methods", "Biological Psychology", "Cognitive Psychology", "Developmental Psychology",
        "Social Psychology", "Abnormal Psychology", "Memory", "Attachment"
    ],
    "History": [
        "World War I & II", "Cold War", "Industrial Revolution", "Ancient Civilizations",
        "Independence Movements", "Civil Rights Movements", "Colonialism", "Revolutions"
    ],
    "Geography": [
        "Climate & Weather", "Population & Migration", "Urbanisation", "Natural Resources",
        "Ecosystems & Biomes", "Rivers & Coasts", "Tectonic Hazards", "Globalisation",
        "Development & Inequality"
    ],
    "ICT": [
        "Hardware & Software", "Word Processing & Spreadsheets", "Databases",
        "Networks & Internet", "Multimedia & Presentations", "E-Safety & Digital Citizenship",
        "Programming Basics", "Data Representation"
    ],
    "Social Studies": [
        "Citizenship & Government", "Geography of Singapore/Region", "History of the Nation",
        "Cultural Diversity", "Economics Basics", "Global Connections"
    ],
    "Social Science": [
        "History", "Geography", "Civics", "Economics", "Disaster Management"
    ],
    "Hindi": [
        "व्याकरण (Grammar)", "गद्य (Prose)", "पद्य (Poetry)", "निबंध लेखन (Essay Writing)",
        "पत्र लेखन (Letter Writing)", "अपठित गद्यांश (Unseen Passage)"
    ],
    "Art": [
        "Drawing & Sketching", "Painting Techniques", "Colour Theory", "Sculpture",
        "Art History & Appreciation", "Digital Art"
    ],
    "Digital Technologies": [
        "Algorithms & Programming", "Data Representation", "Digital Systems",
        "Networks", "Impacts of Technology", "Robotics"
    ],
    "Computing": [
        "Algorithms", "Programming (Python/Scratch)", "Data Representation",
        "Networks", "Computer Systems", "Creating Media", "Online Safety"
    ],
    "Humanities and Social Sciences": [
        "History", "Geography", "Civics & Citizenship", "Economics & Business"
    ],
}

# ─────────────────────────────────────────
#  AUTH FUNCTIONS
# ─────────────────────────────────────────
def sign_up(email, password, name, school):
    try:
        res = supabase.auth.sign_up({"email": email, "password": password})
        if res.user:
            supabase.table("teachers").upsert({
                "id": res.user.id,
                "email": email,
                "name": name,
                "school": school
            }).execute()
            return True, "Account created! You can now log in."
        return False, "Signup failed. Try again."
    except Exception as e:
        return False, str(e)

def sign_in(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        if res.user:
            st.session_state.user          = res.user
            st.session_state.token         = res.session.access_token
            st.session_state.refresh_token = res.session.refresh_token
            # Attach session so RLS policies see auth.uid() correctly
            supabase.auth.set_session(res.session.access_token, res.session.refresh_token)
            return True, "Login successful!"
        return False, "Invalid email or password."
    except Exception as e:
        return False, str(e)

def sign_out():
    supabase.auth.sign_out()
    for key in ["user","token","refresh_token","current_output","current_title",
                "current_task","edit_mode","edit_id","saved_plans"]:
        st.session_state[key] = None if key in ["user","token","refresh_token","edit_id"] else \
                                 [] if key == "saved_plans" else \
                                 False if key == "edit_mode" else ""
    st.rerun()

# ─────────────────────────────────────────
#  DATABASE FUNCTIONS
# ─────────────────────────────────────────
def load_plans():
    try:
        uid = st.session_state.user.id
        res = supabase.table("saved_plans").select("*")\
              .eq("teacher_id", uid).order("created_at", desc=True).execute()
        st.session_state.saved_plans = res.data or []
    except:
        st.session_state.saved_plans = []

def save_plan(title, content, task, curriculum, subject, grade):
    try:
        uid = st.session_state.user.id
        if st.session_state.edit_id:
            supabase.table("saved_plans").update({
                "title": title, "content": content,
                "updated_at": datetime.now().isoformat()
            }).eq("id", st.session_state.edit_id).execute()
            st.success("✅ Updated!")
        else:
            supabase.table("saved_plans").insert({
                "teacher_id": uid, "title": title, "content": content,
                "task": task, "curriculum": curriculum,
                "subject": subject, "grade": grade,
                "created_at": datetime.now().isoformat(),
                "updated_at": datetime.now().isoformat()
            }).execute()
            st.success("✅ Saved!")
        load_plans()
    except Exception as e:
        st.error(f"Save error: {e}")

def delete_plan(plan_id):
    try:
        supabase.table("saved_plans").delete().eq("id", plan_id).execute()
        load_plans()
        st.rerun()
    except Exception as e:
        st.error(f"Delete error: {e}")

# ─────────────────────────────────────────
#  SYLLABUS FUNCTIONS (custom teacher topics)
# ─────────────────────────────────────────
def get_syllabus(curriculum, grade, subject):
    """Return list of topics the teacher has saved for this curriculum/grade/subject, or None."""
    try:
        uid = st.session_state.user.id
        res = supabase.table("syllabus").select("*")\
              .eq("teacher_id", uid)\
              .eq("curriculum", curriculum)\
              .eq("grade", grade)\
              .eq("subject", subject)\
              .execute()
        if res.data:
            topics_str = res.data[0]["topics"]
            return [t.strip() for t in topics_str.split(",") if t.strip()]
        return None
    except Exception:
        return None

def save_syllabus(curriculum, grade, subject, topics_list):
    try:
        uid = st.session_state.user.id
        topics_str = ", ".join(topics_list)
        supabase.table("syllabus").upsert({
            "teacher_id": uid,
            "curriculum": curriculum,
            "grade": grade,
            "subject": subject,
            "topics": topics_str,
            "updated_at": datetime.now().isoformat()
        }, on_conflict="teacher_id,curriculum,grade,subject").execute()
        return True
    except Exception as e:
        st.error(f"Syllabus save error: {e}")
        return False

# ─────────────────────────────────────────
#  AI FUNCTIONS
# ─────────────────────────────────────────
def generate(prompt):
    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": (
                "You are an expert teacher and curriculum designer. "
                "Always produce well-structured, classroom-ready content "
                "with clear headings, bullet points, and timing where relevant."
            )},
            {"role": "user", "content": prompt}
        ],
        max_tokens=2048, temperature=0.7,
    )
    return response.choices[0].message.content

def build_prompt(task, subject, grade, topics, lessons_per_week,
                 duration, curriculum, activity_duration=30, test_duration=45):
    ctx = {
        "IB (International Baccalaureate)":  "Use IB terminology (ATL skills, Learner Profile, Command Terms). Include inquiry questions.",
        "IGCSE (Cambridge)":                  "Use Cambridge command terms (describe, explain, evaluate, analyse). Align with Cambridge assessment objectives.",
        "A-Level (Cambridge)":                "Use higher-order thinking and Cambridge A-Level mark scheme style.",
        "Cambridge Checkpoint":               "Age-appropriate language and Cambridge Checkpoint assessment style.",
        "Singapore (MOE)":                    "Follow Singapore MOE curriculum. Emphasise problem solving and 21CC skills.",
        "CBSE (India)":                       "Follow CBSE curriculum and NCERT guidelines. Align with board exam pattern.",
        "American (Common Core)":            "Follow Common Core State Standards. Include standards codes.",
        "Australian (ACARA)":                 "Follow Australian Curriculum (ACARA). Include strand references.",
        "British (National Curriculum)":      "Follow UK National Curriculum. Use Ofsted-friendly lesson structure.",
    }.get(curriculum, "")
    base = (f"Curriculum: {curriculum}\nSubject: {subject}\nGrade: {grade}\n"
            f"Topics: {topics}\nLessons/week: {lessons_per_week}\n"
            f"Lesson duration: {duration} min\n\n{ctx}\n\n")
    return {
        "📋 Lesson Plan": base + (
            f"Create {lessons_per_week} detailed lesson plans, one per lesson this week. "
            f"Each is {duration} minutes. Label as Lesson 1, Lesson 2, etc. "
            "Each includes: Objectives, Materials, Starter (5 min), Main activities with timings, "
            "Assessment, Plenary, Differentiation, Homework."),
        "🎯 Class Activity": base + (
            f"Design an engaging {activity_duration}-minute classroom activity. "
            "Include: Title, Objective, Step-by-step instructions, Group/individual, "
            "Materials, Expected outcomes, Extension task."),
        "📝 Homework": base + (
            "Create a homework assignment with 5-8 questions (easy/medium/hard). "
            "Include instructions, expected time, learning outcome."),
        "📄 Unit Test": base + (
            "Create a full unit test: 5 MCQ, 5 short answer, 1 essay/long answer. "
            "Show total marks and include mark scheme at end."),
        "📃 Class Test": base + (
            f"Create a {test_duration}-minute class test with 3-5 questions, "
            "mix of types, total marks, and answer key."),
        "❓ Quiz": base + (
            "Create a 10-question quiz with MCQ, true/false, fill-in-the-blank. "
            "Include answers at the end."),
        "📅 Weekly Schedule": base + (
            f"Create a weekly schedule for {lessons_per_week} lessons/week, "
            f"{duration} min each. Include day/time, topic, description, homework slots."),
        "📊 Marking Rubric": base + (
            "Create a marking rubric with 4-5 criteria, 4 performance levels "
            "(Excellent/Good/Satisfactory/Needs Improvement), descriptors, marks."),
    }.get(task, base + f"Create educational content for {task}.")

# ─────────────────────────────────────────
#  EXPORT FUNCTIONS
# ─────────────────────────────────────────
def export_word(title, content):
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}")
    doc.add_paragraph("")
    for line in content.split("\n"):
        if   line.startswith("# "):   doc.add_heading(line[2:], level=1)
        elif line.startswith("## "):  doc.add_heading(line[3:], level=2)
        elif line.startswith("### "): doc.add_heading(line[4:], level=3)
        elif line.startswith(("- ","* ")): doc.add_paragraph(line[2:], style="List Bullet")
        elif line.strip(): doc.add_paragraph(line)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

def export_excel(title, content):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Lesson Plan"
    hf = PatternFill("solid", fgColor="2E4057")
    ws.merge_cells("A1:C1"); ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = hf; ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:C2")
    ws["A2"] = f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}"
    ws["A2"].alignment = Alignment(horizontal="center")
    for col, h in enumerate(["Content","Details","Notes"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(color="FFFFFF", bold=True); c.fill = hf
        c.alignment = Alignment(horizontal="center")
    row = 4
    for line in content.split("\n"):
        if line.strip():
            ws.cell(row=row, column=1, value=line.strip())
            ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
            row += 1
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 30
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═════════════════════════════════════════
#  LOGIN / SIGNUP PAGE
# ═════════════════════════════════════════
if not st.session_state.user:
    st.title("📚 AI Teacher Assistant")
    st.caption("Smart lesson planning powered by AI — Free for teachers")
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        mode = st.radio("", ["🔑 Login", "📝 Sign Up"], horizontal=True, label_visibility="collapsed")

        if mode == "🔑 Login":
            st.subheader("Welcome back!")
            email    = st.text_input("Email", placeholder="teacher@school.com")
            password = st.text_input("Password", type="password")
            if st.button("Login", type="primary", use_container_width=True):
                if email and password:
                    ok, msg = sign_in(email, password)
                    if ok:
                        load_plans()
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Please enter email and password.")

        else:
            st.subheader("Create your account")
            name     = st.text_input("Your Name",    placeholder="e.g. Mr. Akhilesh")
            school   = st.text_input("School Name",  placeholder="e.g. Singapore International School")
            email    = st.text_input("Email",         placeholder="teacher@school.com")
            password = st.text_input("Password",      type="password", help="Minimum 6 characters")
            password2= st.text_input("Confirm Password", type="password")
            if st.button("Create Account", type="primary", use_container_width=True):
                if not all([name, school, email, password, password2]):
                    st.warning("Please fill in all fields.")
                elif password != password2:
                    st.error("Passwords do not match.")
                elif len(password) < 6:
                    st.error("Password must be at least 6 characters.")
                else:
                    ok, msg = sign_up(email, password, name, school)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

    st.stop()

# ═════════════════════════════════════════
#  MAIN APP (logged in)
# ═════════════════════════════════════════
st.title("📚 AI Teacher Assistant")
st.caption("Developed by Akhilesh Kumar Srivastava — AI Teaching Assistant for IB, IGCSE, Singapore, A-Level & more")

with st.expander("ℹ️ About this tool / How to use"):
    st.markdown("""
    **Purpose:** Save time on lesson planning, activities, homework, tests, and quizzes — all generated instantly by AI.

    **Who can use it:** Any teacher — just sign up with your email to get your own private workspace.

    **How to use:**
    1. Select your Curriculum, Grade, Subject, and Topics in the sidebar
    2. Set lessons per week and lesson duration
    3. Choose what to generate (Lesson Plan, Quiz, Test, etc.)
    4. Click ⚡ Generate
    5. Save to Cloud, Edit, or Download as Word / Excel

    **Developed by:** Akhilesh Kumar Srivastava
    """)

# ── Sidebar ───────────────────────────────
with st.sidebar:
    # Teacher info
    st.success(f"👋 {st.session_state.user.email}")
    if st.button("🚪 Logout", use_container_width=True):
        sign_out()

    st.divider()
    st.header("⚙️ Settings")

    curriculum = st.selectbox("🎓 Curriculum", list(GRADE_OPTIONS.keys()))
    grade      = st.selectbox("📊 Grade / Year", GRADE_OPTIONS[curriculum])

    subject_list = SUBJECT_OPTIONS.get(curriculum, ["Computer Science", "Mathematics", "Science", "English"])
    subject_list = subject_list + ["Other (type manually)"]
    subject_choice = st.selectbox("📖 Subject", subject_list)

    if subject_choice == "Other (type manually)":
        subject = st.text_input("Enter subject name", placeholder="e.g. Computer Science")
    else:
        subject = subject_choice

    # ── Topics: use custom syllabus if teacher has saved one, else built-in list ──
    custom_topics = get_syllabus(curriculum, grade, subject)
    default_list  = TOPIC_OPTIONS.get(subject, [])
    topic_list    = custom_topics if custom_topics is not None else default_list

    if topic_list:
        selected_topics = st.multiselect(
            "📌 Topics (select one or more)",
            topic_list,
            help="Select topics to include — you can also add extra topics below"
        )
    else:
        selected_topics = []
        st.info("📋 No topic list yet for this subject/grade. Add your syllabus below, or type topics manually.")

    extra_topics = st.text_area(
        "➕ Additional topics (optional)",
        placeholder="Type any extra topics not in the list above, separated by commas",
        height=60
    )

    topic_parts = selected_topics + (
        [t.strip() for t in extra_topics.split(",") if t.strip()] if extra_topics else []
    )
    topics = ", ".join(topic_parts)

    # ── Syllabus upload / edit ──
    with st.expander("📚 My Syllabus (add / edit topics for this subject)"):
        st.caption(f"For: {curriculum} · {grade} · {subject}")
        if custom_topics:
            st.success(f"✅ You have {len(custom_topics)} custom topics saved for this subject.")
            existing_text = ", ".join(custom_topics)
        else:
            st.caption("No custom syllabus saved yet. Add your topics below (comma separated) — "
                       "these will appear in the dropdown above for future use.")
            existing_text = ", ".join(default_list)

        syllabus_text = st.text_area(
            "Topics (comma separated)",
            value=existing_text,
            height=120,
            key=f"syllabus_{curriculum}_{grade}_{subject}"
        )
        if st.button("💾 Save Syllabus", use_container_width=True,
                      key=f"save_syllabus_{curriculum}_{grade}_{subject}"):
            new_topics = [t.strip() for t in syllabus_text.split(",") if t.strip()]
            if new_topics:
                if save_syllabus(curriculum, grade, subject, new_topics):
                    st.success("✅ Syllabus saved! It will now appear in the Topics dropdown.")
                    st.rerun()
            else:
                st.warning("Please enter at least one topic.")

    lessons_per_week = st.number_input("Lessons per week", min_value=1, max_value=10, value=3)
    duration         = st.number_input("Lesson duration (min)", min_value=20, max_value=180, value=45)

    st.divider()
    task = st.selectbox("🛠 What to generate", [
        "📋 Lesson Plan", "🎯 Class Activity", "📝 Homework",
        "📄 Unit Test", "📃 Class Test", "❓ Quiz",
        "📅 Weekly Schedule", "📊 Marking Rubric",
    ])

    activity_duration = duration
    test_duration     = 45

    if task == "🎯 Class Activity":
        st.info("⏱ Activity duration?")
        activity_duration = st.number_input("Activity duration (min)", min_value=5, max_value=180, value=30)
    elif task == "📃 Class Test":
        st.info("⏱ Class test duration?")
        test_duration = st.number_input("Test duration (min)", min_value=10, max_value=180, value=45)
    elif task == "📋 Lesson Plan":
        st.info(f"📅 Will generate {int(lessons_per_week)} lesson plan(s)")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        generate_btn = st.button("⚡ Generate", type="primary", use_container_width=True)
    with col2:
        if st.button("✖ Cancel", use_container_width=True):
            st.session_state.current_output = ""
            st.session_state.current_title  = ""
            st.session_state.edit_mode      = False
            st.session_state.edit_id        = None
            st.rerun()

    if st.button("🆕 New", use_container_width=True):
        st.session_state.current_output = ""
        st.session_state.current_title  = ""
        st.session_state.edit_mode      = False
        st.session_state.edit_id        = None
        st.rerun()

    st.divider()
    st.subheader("📂 My Saved Plans")
    if st.button("🔄 Refresh", use_container_width=True):
        load_plans()

    plans = st.session_state.saved_plans
    if plans:
        for plan in plans:
            col_a, col_b = st.columns([4, 1])
            with col_a:
                label = plan["title"][:25] + "…" if len(plan["title"]) > 25 else plan["title"]
                if st.button(f"📄 {label}", key=f"open_{plan['id']}", use_container_width=True):
                    st.session_state.current_output = plan["content"]
                    st.session_state.current_title  = plan["title"]
                    st.session_state.current_task   = plan.get("task", "")
                    st.session_state.edit_id        = plan["id"]
                    st.session_state.edit_mode      = False
                    st.rerun()
            with col_b:
                if st.button("🗑", key=f"del_{plan['id']}"):
                    delete_plan(plan["id"])
    else:
        st.caption("No saved plans yet. Click 🔄 Refresh after saving, or save your work using 💾 Save to Cloud on the right before leaving.")

# ── Main area ─────────────────────────────
col_main, col_actions = st.columns([3, 1])

with col_main:
    if generate_btn:
        if not subject or not topics:
            st.warning("Please fill in Subject and Topics.")
        else:
            title = f"{task} — {curriculum} · {subject} · {grade}"
            with st.spinner("Generating... ⏳"):
                try:
                    result = generate(build_prompt(
                        task, subject, grade, topics,
                        lessons_per_week, duration, curriculum,
                        activity_duration, test_duration
                    ))
                    st.session_state.current_output = result
                    st.session_state.current_title  = title
                    st.session_state.current_task   = task
                    st.session_state.edit_mode      = False
                    st.session_state.edit_id        = None
                except Exception as e:
                    st.error(f"Error: {e}")

    if st.session_state.current_output:
        st.subheader(st.session_state.current_title)
        if st.session_state.edit_mode:
            edited = st.text_area("✏️ Edit:", value=st.session_state.current_output, height=500)
            c1, c2 = st.columns(2)
            with c1:
                if st.button("✅ Save edits", type="primary"):
                    st.session_state.current_output = edited
                    if st.session_state.edit_id:
                        save_plan(
                            st.session_state.current_title, edited,
                            st.session_state.current_task, "", "", ""
                        )
                    st.session_state.edit_mode = False
                    st.rerun()
            with c2:
                if st.button("❌ Cancel edit"):
                    st.session_state.edit_mode = False
                    st.rerun()
        else:
            st.markdown(st.session_state.current_output)
    else:
        st.info("👈 Fill in your details on the left and click ⚡ Generate")

with col_actions:
    if st.session_state.current_output:
        st.subheader("Actions")

        if st.button("💾 Save to Cloud", use_container_width=True):
            save_plan(
                st.session_state.current_title,
                st.session_state.current_output,
                st.session_state.current_task,
                curriculum, subject, grade
            )

        if st.button("✏️ Edit", use_container_width=True):
            st.session_state.edit_mode = True
            st.rerun()

        st.divider()

        word_buf = export_word(st.session_state.current_title, st.session_state.current_output)
        st.download_button(
            "⬇️ Download Word", data=word_buf,
            file_name=f"{st.session_state.current_title[:40]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

        if "Lesson Plan" in st.session_state.current_task:
            xl_buf = export_excel(st.session_state.current_title, st.session_state.current_output)
            st.download_button(
                "📊 Download Excel", data=xl_buf,
                file_name=f"{st.session_state.current_title[:40]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.download_button(
            "📄 Download .txt", data=st.session_state.current_output,
            file_name=f"{st.session_state.current_title[:40]}.txt",
            mime="text/plain", use_container_width=True
        )

        st.divider()
        if st.button("🗑 Clear", use_container_width=True):
            st.session_state.current_output = ""
            st.session_state.current_title  = ""
            st.session_state.edit_mode      = False
            st.session_state.edit_id        = None
            st.rerun()