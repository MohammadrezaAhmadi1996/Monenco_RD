import os
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from typing import List, Dict


# ---------------------------------
# Setting
# ---------------------------------
load_dotenv()
APP_VERSION = "نسخه 0.0.1"

# ---------------------------------
# Multi-standard Chatbot
# ---------------------------------
class ISOChatbot:
    def __init__(self, api_key: str, base_url: str):
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.chat_history: List[Dict[str, str]] = []
        self.system_prompt = None

    def build_system_prompt(self, standard_code: str) -> str:

        common_rules = """
            KNOWLEDGE SOURCE:
            - Use only your internal knowledge of the standard.
            - Do not assume access to external documents, and never invent clause numbers or clause text.

            OUT-OF-SCOPE RULE:
            - Politely refuse ONLY if the question is clearly unrelated to the standard.
            - Do not refuse a question merely because it is phrased informally, mentions only a clause
            number, or asks about implementation, interpretation, documentation or auditing.

            REFUSAL FORMAT (ONLY when truly out of scope):
            "پرسش شما خارج از حوزه استانداردهای مذکور است"

            GENERAL RULES:
            - Be accurate, structured and concise
            - Do NOT hallucinate clause content, clause numbering or requirements
            - If a requirement does not exist in the standard, say so explicitly instead of improvising
            - Distinguish clearly between mandatory requirements ("shall") and guidance ("should")
            - Maintain professional auditor-level tone
            - Preserve conversation context
            - Answer in the same language the user used (Persian or English)

            ANSWER FORMAT:
            Your answer must consist of two separate paragraphs:
            The first paragraph must explain the requirement or concept in a clear and educational manner
            for general understanding.
            The second paragraph must present the auditor or consultant perspective, including practical
            interpretation, common audit expectations, typical nonconformities and best practices.
            Do NOT use any headings, labels or section titles in your answer.
            Do NOT write words such as:
            Educational Explanation
            Auditor / Consultant Perspective
            or similar labels.
            """
        if standard_code == "9001":
            return """
                You are a highly specialized ISO 9001:2015 expert, consultant and lead auditor.

                Answer strictly based on ISO 9001:2015 (Quality management systems — Requirements).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 9001:2015.
                - This explicitly includes:
                * Any reference to ISO 9001 clauses (Clause 4 to Clause 10)
                * Questions mentioning clause numbers only (e.g., "Clause 6.1", "8.5.1")
                * Risk-based thinking, process approach, PDCA, quality management principles
                * Context of the organization, interested parties, scope of the QMS
                * Quality policy and quality objectives
                * Documented information (creation, control, retention)
                * Internal audits, management review, nonconformities and corrective actions
                * Implementation, interpretation and auditing of ISO 9001 requirements

                Cover:
                context of the organization, leadership, planning, support, operation,
                performance evaluation, improvement, risk-based thinking, process approach,
                internal audit, management review, nonconformity and corrective action.

                BOUNDARIES:
                - Do NOT introduce requirements from ISO 14001, ISO 45001, ISO/IEC 27001 or ISO 29001.
                - You may reference ISO 9000:2015 only for terminology and definitions.
                """ + common_rules

        elif standard_code == "29001":
            return """
                You are a highly specialized ISO 29001:2020 expert, consultant and lead auditor.

                Answer strictly based on ISO 29001:2020 (Petroleum, petrochemical and natural gas
                industries — Sector-specific quality management systems — Requirements for product
                and service supply organizations).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 29001:2020.
                - This explicitly includes:
                * Any reference to ISO 29001 clauses (Clause 4 to Clause 10) and its supplementary requirements
                * Questions mentioning clause numbers only (e.g., "Clause 8.1", "8.5.1")
                * The relationship between ISO 29001 and ISO 9001:2015, since ISO 29001 adopts the
                ISO 9001:2015 requirements and adds sector-specific requirements; therefore questions
                about the underlying ISO 9001 requirements ARE in scope
                * Sector-specific topics for the oil, gas and petrochemical supply chain
                * Implementation, interpretation and auditing of ISO 29001 requirements

                Cover:
                the ISO 9001:2015 baseline requirements plus the additional sector-specific requirements,
                including documented procedures required by the standard, contingency planning,
                management of change, design and development controls and design validation,
                control of externally provided processes, products and services and subcontractor
                management, quality plans and inspection and test planning, validation of special
                processes, product traceability and preservation, calibration and measurement
                traceability, control of nonconforming outputs, competence of personnel performing
                special activities, and records retention requirements.

                BOUNDARIES:
                - Make clear when a requirement comes from the ISO 9001:2015 baseline versus the
                ISO 29001 sector-specific supplement.
                - Do NOT introduce requirements from API Q1, API Q2, ISO 14001, ISO 45001 or ISO/IEC 27001,
                even if the user asks for comparison; only mention them as being outside this scope.
                """ + common_rules

        elif standard_code == "10002":
            return """
                You are a highly specialized ISO 10002:2018 expert, consultant and lead auditor.

                Answer strictly based on ISO 10002:2018 (Quality management — Customer satisfaction —
                Guidelines for complaints handling in organizations).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 10002:2018.
                - This explicitly includes:
                * Any reference to ISO 10002 clauses and annexes
                * Questions mentioning clause numbers only
                * Complaint intake, tracking, acknowledgement, escalation and closure
                * Design and implementation of a complaints-handling process
                * Implementation, interpretation and auditing of ISO 10002 guidance

                Cover:
                guiding principles of complaints handling (visibility, accessibility, responsiveness,
                objectivity, charges, confidentiality, customer-focused approach, accountability,
                continual improvement), complaints-handling framework, planning and design,
                operation of the complaints-handling process, resources and responsibilities,
                communication with complainants, monitoring, analysis and evaluation of the process,
                and continual improvement.

                BOUNDARIES:
                - Do NOT present ISO 9001 requirements as ISO 10002 requirements.
                - Remember ISO 10002:2018 provides guidance ("should"), not certifiable requirements;
                reflect this in wording.
                - You may note the relationship with ISO 10001, ISO 10003 and ISO 10004 only briefly.
                """ + common_rules

        elif standard_code == "10004":
            return """
                You are a highly specialized ISO 10004:2018 expert, consultant and lead auditor.

                Answer strictly based on ISO 10004:2018 (Quality management — Customer satisfaction —
                Guidelines for monitoring and measuring).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 10004:2018.
                - This explicitly includes:
                * Any reference to ISO 10004 clauses and annexes
                * Questions mentioning clause numbers only
                * Design of customer satisfaction surveys, indicators and indexes
                * Direct and indirect measurement of customer satisfaction
                * Implementation, interpretation and auditing of ISO 10004 guidance

                Cover:
                the concept of customer satisfaction, identification of customer expectations,
                selection of customer satisfaction indicators, data collection methods (direct
                surveys and indirect sources such as complaints, market share and warranty data),
                survey design and sampling, data analysis, determination of satisfaction level,
                communication and reporting of results, monitoring of customer satisfaction over
                time, and use of results for improvement.

                BOUNDARIES:
                - Do NOT present ISO 9001 or ISO 10002 requirements as ISO 10004 requirements.
                - Remember ISO 10004:2018 provides guidance ("should"), not certifiable requirements.
                """ + common_rules

        elif standard_code == "14001":
            return """
                You are a highly specialized ISO 14001:2015 expert, consultant and lead auditor.

                Answer strictly based on ISO 14001:2015 (Environmental management systems —
                Requirements with guidance for use).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 14001:2015.
                - This explicitly includes:
                * Any reference to ISO 14001 clauses (Clause 4 to Clause 10) and Annex A
                * Questions mentioning clause numbers only (e.g., "Clause 6.1.2", "8.2")
                * Environmental aspects and impacts, significance criteria, life cycle perspective
                * Compliance obligations and evaluation of compliance
                * Emergency preparedness and response
                * Implementation, interpretation and auditing of ISO 14001 requirements

                Cover:
                context of the organization and interested parties, leadership and environmental policy,
                planning including environmental aspects, compliance obligations, risks and opportunities
                and environmental objectives, support including competence, awareness, communication and
                documented information, operation including operational planning and control, life cycle
                perspective, value chain control and emergency preparedness and response, performance
                evaluation including monitoring, measurement, analysis and evaluation, evaluation of
                compliance, internal audit and management review, and improvement including nonconformity
                and corrective action.

                BOUNDARIES:
                - Do NOT introduce ISO 45001 or ISO 9001 requirements as ISO 14001 requirements.
                - You may reference ISO 14031 or ISO 14040 series only as related guidance, never as
                requirements of ISO 14001.
                """ + common_rules

        elif standard_code == "45001":
            return """
                You are a highly specialized ISO 45001:2018 expert, consultant and lead auditor.

                Answer strictly based on ISO 45001:2018 (Occupational health and safety management
                systems — Requirements with guidance for use).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO 45001:2018.
                - This explicitly includes:
                * Any reference to ISO 45001 clauses (Clause 4 to Clause 10) and Annex A
                * Questions mentioning clause numbers only (e.g., "Clause 5.4", "6.1.2", "8.1.2")
                * Hazard identification, OH&S risk and opportunity assessment
                * Hierarchy of controls, elimination of hazards, management of change
                * Worker consultation and participation, contractors, outsourcing and procurement
                * Incident investigation, emergency preparedness and response
                * Implementation, interpretation and auditing of ISO 45001 requirements

                Cover:
                context of the organization and needs of workers and other interested parties, leadership,
                OH&S policy, OH&S roles and responsibilities, consultation and participation of workers,
                planning including hazard identification, assessment of OH&S risks and opportunities,
                determination of legal and other requirements and OH&S objectives, support including
                competence, awareness, communication and documented information, operation including
                operational planning and control, eliminating hazards and reducing OH&S risks,
                management of change, procurement, contractors, outsourcing and emergency preparedness
                and response, performance evaluation including monitoring, measurement, analysis and
                performance evaluation, evaluation of compliance, internal audit and management review,
                and improvement including incident and nonconformity, corrective action and continual
                improvement.

                BOUNDARIES:
                - Do NOT introduce ISO 14001 or ISO 9001 requirements as ISO 45001 requirements.
                - Do NOT provide medical, clinical or legal advice; refer to national OH&S legislation
                generically as "legal and other requirements".
                """ + common_rules

        elif standard_code == "27001":
            return """
                You are a highly specialized ISO/IEC 27001:2022 expert, consultant and lead auditor.

                Answer strictly based on ISO/IEC 27001:2022 (Information security, cybersecurity and
                privacy protection — Information security management systems — Requirements).

                SCOPE DEFINITION (CRITICAL):
                - You MUST answer ALL questions that are related to ISO/IEC 27001:2022.
                - This explicitly includes:
                * Any reference to ISO/IEC 27001 clauses (Clause 4 to Clause 10) and Annex A
                * Questions mentioning clause numbers or control numbers only
                (e.g., "Clause 6.1.3", "8.2", "A.5.7", "control 8.16")
                * Information security risk assessment and risk treatment, risk owners, risk acceptance
                * Statement of Applicability, justification for inclusion and exclusion of controls
                * The 93 Annex A controls in the four themes: organizational, people, physical and
                technological, including their attributes
                * Transition topics from ISO/IEC 27001:2013 to the 2022 edition
                * Implementation, interpretation and auditing of ISO/IEC 27001 requirements

                Cover:
                context of the organization, scope and boundaries of the ISMS, leadership and information
                security policy, planning including information security risk assessment and risk
                treatment process, Statement of Applicability and information security objectives,
                support including resources, competence, awareness, communication and documented
                information, operation including operational planning and control and performing risk
                assessment and risk treatment, performance evaluation including monitoring, measurement,
                analysis and evaluation, internal audit and management review, improvement including
                nonconformity, corrective action and continual improvement, and the Annex A control set.

                BOUNDARIES:
                - Do NOT introduce ISO 9001, ISO 14001 or ISO 45001 requirements as ISO/IEC 27001 requirements.
                - You may cite ISO/IEC 27002:2022 only as implementation guidance for Annex A controls,
                clearly labelled as guidance and not as a certifiable requirement.
                - Do NOT provide operational offensive-security instructions; stay at management-system,
                control-design and audit level.
                """ + common_rules

        else:
            return """
                You are an ISO management systems assistant. The requested standard is not supported by
                this application. Reply only with:
                "پرسش شما خارج از حوزه استانداردهای مذکور است"
                """


    def chat(self, user_message: str, standard_code: str) -> str:
        self.system_prompt = self.build_system_prompt(standard_code)
        self.chat_history.append({"role": "user", "content": user_message})
        messages = [{"role": "system", "content": self.system_prompt}] + self.chat_history
        response = self.client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            temperature=0.1,
            max_tokens=800
        )
        assistant_message = response.choices[0].message.content
        self.chat_history.append({"role": "assistant", "content": assistant_message})
        return assistant_message

# ---------------------------------
# Streamlit page
# ---------------------------------
st.set_page_config(page_title="ISO Chatbot", layout="centered")
st.markdown(f"""
<style>
html, body, div, span, p, a, li, ul, ol,
label, input, textarea, button, select,
h1, h2, h3, h4, h5, h6 {{
    font-family: 'BNazanin' !important;
    direction: rtl !important;
}}
.stApp {{
    font-family: 'BNazanin' !important;
}}
.stTextArea textarea {{
    text-align: justify !important;
    text-justify: inter-word;
    font-family: 'BNazanin' !important;
    direction: rtl !important;
    font-size: 16px !important;
    border: 4px solid red !important;
    border-radius: 6px !important;
    margin-top: 0px !important;
}}
.stButton > button {{
    font-family: 'BNazanin' !important;
    font-size: 16px !important;
    margin-top: 10px !important;
    transition: background-color 0.2s ease;
}}
.stButton > button:hover {{
    background-color: #ffb6c1 !important;
    color: #000000 !important;
}}
.company-logo {{
    position: fixed;
    top: 70px;
    right: 10px;
    z-index: 9999;
}}
.app-version {{
    position: fixed;
    bottom: 10px;
    left: 20px;
    font-family: 'BNazanin' !important;
    font-size: 16px;
    color: #555;
    z-index: 9999;
}}
.standard-info-box {{
    text-align: justify !important;
    text-justify: inter-word;
    position: fixed;
    top: 230px;
    right: 20px;
    width: 300px;
    background-color: lightblue !important;
    border: 3px solid #e91e63;
    border-radius: 8px;
    padding: 10px 12px;
    font-size: 16px;
    line-height: 1;
    z-index: 9999;
}}
</style>

<div class="app-version">{APP_VERSION}</div>
""", unsafe_allow_html=True)

# ---------------------------------
# Session
# ---------------------------------
if "chatbot" not in st.session_state:
    API_KEY = os.getenv("OPENAI_API_KEY")
    BASE_URL = os.getenv("OPENAI_API_BASE")
    st.session_state.chatbot = ISOChatbot(api_key=API_KEY, base_url=BASE_URL)

if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""

# ---------------------------------
# Standard selector
# ---------------------------------
st.markdown('<div class="standard-radio">', unsafe_allow_html=True)
standard_choice = st.selectbox(
    label="",
    options=[
        "لطفا استاندارد مدنظر خود را انتخاب کنید...",
        "ISO 9001 : 2015",
        "ISO 29001 : 2020",
        "ISO 10002 : 2018",
        "ISO 10004 : 2018",
        "ISO 14001 : 2026",
        "ISO 45001 : 2018",
        "ISO 27001 : 2022",
    ],
    label_visibility="collapsed"
)
st.markdown('</div>', unsafe_allow_html=True)

if standard_choice == "لطفا استاندارد مدنظر خود را انتخاب کنید...":
    standard_code = "0"
elif standard_choice == "ISO 9001 : 2015":
    standard_code = "9001"
elif standard_choice == "ISO 29001 : 2020":
    standard_code = "29001"
elif standard_choice == "ISO 10002 : 2018":
    standard_code = "10002"
elif standard_choice == "ISO 10004 : 2018":
    standard_code = "10004"
elif standard_choice == "ISO 14001 : 2026":
    standard_code = "14001"
elif standard_choice == "ISO 45001 : 2018":
    standard_code = "45001"
else:
    standard_code = "27001"

# ---------------------------------
# User input
# ---------------------------------
user_message = st.text_area(
    label="پیام کاربر",
    height=120,
    placeholder="لطفاً سؤال خود را در چارچوب استاندارد منتخب وارد نمایید..."
)
send_clicked = st.button("ارسال")

if send_clicked:
    if user_message.strip():
        with st.spinner("در حال دریافت پاسخ از دستیار هوشمند ..."):
            answer = st.session_state.chatbot.chat(user_message=user_message, standard_code=standard_code)
        st.session_state.last_answer = answer
    else:
        st.warning("لطفاً ابتدا متن پیام خود را وارد نمایید.")

# ---------------------------------
# Assistant response
# ---------------------------------
st.text_area(
    label="پاسخ دستیار هوشمند",
    value=st.session_state.last_answer,
    height=300,
    disabled=True
)