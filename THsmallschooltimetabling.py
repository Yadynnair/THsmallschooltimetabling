import streamlit as st
import numpy as np
import pandas as pd
import altair as alt
import pulp as pu
import xlsxwriter

st.set_page_config(
    page_title="จัดตารางเรียนตารางสอนสำหรับโรงเรียนขนาดเล็ก",
    page_icon="📚",
)

with st.sidebar:
    st.header('โปรแกรมนี้จัดทำโดย')
    st.markdown('''
        ดร.ศรณย์เศรษฐ์ โสกันธิกา

        วิทยาลัยนานาชาตินวัตกรรมดิจิทัล มหาวิทยาลัยเชียงใหม่ :elephant:
        
        สามารถติดต่อได้ผ่านทางอีเมล s.sokantika@gmail.com
    ''')

    with st.expander(''):
        st.header('เรื่อยเปื่อย')
        st.markdown(''' 
            ความคิดที่จัดทำโปรแกรมนี้เกิดมาจากการที่ต้องการแก้ปัญหาให้ภรรยาซึ่งเป็นครูในโรงเรียนขนาดเล็ก 
            ผู้จัดทำเลยมีความคิดว่า แทนที่จะช่วยแก้ปัญหาให้ภรรยาตัวเองคนเดียวทำไมไม่ช่วยแก้ปัญหาให้คุณครูคนอื่นไปด้วยเลย สุดท้ายเลยกลายเป็นโปรเจคนี้ออกมาครับ
            
            ถ้ามีข้อเสนอแนะเพิ่มเติมสามารถส่งมาทางอีเมลได้เลยครับ
        ''')

st.title('ยินดีต้อนรับเข้าสู่โปรแกรมช่วยจัดตารางสอนสำหรับโรงเรียนขนาดเล็ก')
st.markdown('''
    โปรแกรมนี้เป็นต้นแบบที่จัดทำขึ้นโดยคาดหวังที่จะแบ่งเบาภาระงานของครูในโรงเรียนขนาดเล็กในประเทศไทย
    เนื่องจากโรงเรียนอาจมีทรัพยากรไม่เพียงพอจึงไม่สามารถจัดซื้อโปรแกรมในการช่วยจัดตารางสอน 
    ทำให้คุณครูต้องเสียเวลาในการจัดตารางเรียนตารางสอนและอาจเกิดความผิดพลาดได้ง่าย
''')

st.header("อัพโหลดข้อมูล")

st.write("หลังจากที่คุณครูได้ประชุมเพื่อตกลงภาระงานสอนแล้ว **โปรดคลิ๊กปุ่มด้านล่าง**เพื่อดาวน์โหลดแม่แบบสำหรับกรอกข้อมูล")

with open('โรงเรียนบ้านห่างไกล_2_2565.xlsx', 'rb') as my_file:
    st.download_button(
        label = '📥 ดาวน์โหลดแม่แบบ 📥', 
        data = my_file, 
        file_name = 'โรงเรียนบ้านห่างไกล_2_2565.xlsx', 
        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')      

st.write("หลังจากที่ท่านได้ดาวน์โหลดแม่แบบแล้ว") 
st.write("  1. โปรดเปลี่ยนชื่อไฟล์เป็น **ชื่อโรงเรียน_เทอม_ปีการศึกษา** เช่น โรงเรียนบ้านห่างไกล_2_2565")
st.write("  2. โปรดเปลี่ยนชื่อแต่ละชีทเป็น **ชื่อ นามสกุล** ของครูแต่ละท่าน เช่น สมหมาย รักสอน")
st.write("  3. เลขที่ปรากฎในตารางหมายถึงจำนวนคาบที่สอนในวิชาและชั้นเรียนนั้น ๆ ในหนึ่งสัปดาห์ เช่น ในชีทแรก ครูสมหมาย สอนวิชาภาษาไทย ป.4 จำนวน 4 คาบ ใน 1 สัปดาห์")
st.write("  4. ถ้าวิชาพละศึกษาและวิชาสุขศึกษาใช้คาบเรียนร่วมกัน ให้ใส่เลขในช่องของวิชาพละศึกษาอย่างเดียว")

uploaded_file = st.file_uploader("อัพโหลดแม่แบบ")
# upload_status = False
TA = {}

if uploaded_file is not None:
    # upload_status = True
    df = pd.read_excel(uploaded_file,sheet_name=None, engine='openpyxl')
    
    school_name = uploaded_file.name.split('_')[0]
    semester = uploaded_file.name.split('_')[1]
    year = uploaded_file.name.split('_')[2]
    
# Create database

    Teacher = [*df] #get teacher list
    
    Grades = []
    Remove_grades = ['วิชา/ชั้น','อนุบาล']
    for key in df[Teacher[0]].columns:
        if key not in Remove_grades:
            Grades.append(key)
    
    Allgrades = [g for g in df[Teacher[0]].columns if g != 'วิชา/ชั้น']

    Subjects = []
    Remove_subjects = ['ประจำชั้น']
    for key in df[Teacher[0]]['วิชา/ชั้น']:
            if key not in Remove_subjects:
                Subjects.append(key)

    for t in Teacher: #Create Teacher assignment
        Dummy = {}
        for g in Grades:
            Dummy2 = []
            for s in range(len(Subjects)):
                Dummy2.append(round(df[t][g][s]) if pd.notna(df[t][g][s]) else 0)
            Dummy[g] = Dummy2
        
        Dummy['ประจำชั้น'] = ''
        for g in Grades+['อนุบาล']:
            if df[t][g][df[t][df[t]['วิชา/ชั้น'] == 'ประจำชั้น'].index[0]] == 1:
                Dummy['ประจำชั้น'] = g

        TA[t] = Dummy

    # Actual Classes
    Classes = []
    for t in Teacher:
        for j in range(len(Subjects)):
            for i in Grades:
                if TA[t][i][j] !=0:
                    for k in range(TA[t][i][j]):
                        Classes.append((t,i,Subjects[j],k))

    # teaching workload 
    teaching_workload = {t:[] for t in Teacher}
    for t in Teacher:
        for c in Classes:
            if c[0] == t:
                teaching_workload[t].append(c)

    # student plan
    student_plan = {g:[] for g in Grades}
    for g in Grades:
        for c in Classes:
            if c[1] == g:
                student_plan[g].append(c)

    # subject plan
    subject_plan = {s:[] for s in Subjects}
    for s in Subjects:
        for c in Classes:
            if c[2] == s:
                subject_plan[s].append(c)

# Upload data not too large

    Teacher_limit = 20
    Grades_limit = 15
    Subjects_limit = 15
    error_counter = False

    if len(Teacher) > Teacher_limit:
        st.error("จำนวนครูผู้สอนมากเกินไป จำกัดอยู่ที่ {} คน".format(Teacher_limit))
        error_counter = True

    if len(Grades) > Grades_limit:
        st.error("จำนวนชั้นเรียนมีมากเกินไป จำกัดอยู่ที่ {} ชั้นเรียน".format(Grades_limit))
        error_counter = True

    if len(Subjects) > Subjects_limit:
        st.error("จำนวนวิชามีมากเกินไป จำกัดอยู่ที่ {} วิชา".format(Subjects_limit))
        error_counter = True

# Visualize upload data

    with st.expander("โปรดตรวจสอบภาพรวมของข้อมูลที่อัพโหลด"):
        #Teacher Chart
        t_chart_data = pd.DataFrame({
                'ครู': Teacher,
                'คาบ/สัปดาห์': [sum(TA[t][g][Subjects.index(s)] for g in Grades for s in Subjects) for t in Teacher]
                })

        teacher_chart = alt.Chart(t_chart_data).mark_bar().encode(x=alt.X('ครู',axis=alt.Axis(title=' '), sort='-y'), y=alt.Y('คาบ/สัปดาห์',axis=alt.Axis(tickCount=20))).properties(title='จำนวนคาบที่สอน')
      
        st.altair_chart(teacher_chart, use_container_width=True)


        #Student Chart
        s_chart_data = pd.DataFrame({
                'นักเรียน': Grades,
                'คาบ/สัปดาห์': [sum(TA[t][g][Subjects.index(s)] for t in Teacher for s in Subjects) for t in Grades]
                })

        student_chart = alt.Chart(s_chart_data).mark_bar().encode(x=alt.X('นักเรียน',axis=alt.Axis(title=' ')),y=alt.Y('คาบ/สัปดาห์',axis=alt.Axis(tickCount=25))).properties(title='จำนวนคาบที่เรียน')

        st.altair_chart(student_chart, use_container_width=True)

        #display assigned workload
        st.write('ภาระงานสอนที่ได้รับมอบหมาย') 
        df_show = pd.DataFrame(columns=Grades)
        for s in Subjects:
            Dummy = []
            for g in Grades:
                TDummy = [ t for t in Teacher if TA[t][g][Subjects.index(s)] != 0]
                DDummy = ''
                if  TDummy is not np.empty:
                    for t in TDummy:
                        DDummy = DDummy + 'ครู{} {} '.format(t.split(' ')[0],round(TA[t][g][Subjects.index(s)]))                
                Dummy.append(DDummy)       
            df_show.loc[s] = Dummy

        st.dataframe(df_show)

        #display homeroom teacher
        st.write('ครูประจำชั้น')
        df_show = pd.DataFrame(columns=Allgrades)
        Dummy = []
        for g in Allgrades:
            TDummy = [ t for t in Teacher if TA[t]['ประจำชั้น'] == "{}".format(g)]
            DDummy = ''
            if  TDummy is not np.empty:
                for t in TDummy:
                    DDummy = DDummy + 'ครู{}'.format(t.split(' ')[0])                
            Dummy.append(DDummy)       
        df_show.loc['ครูประจำชั้น'] = Dummy

        st.dataframe(df_show)

    if error_counter == False:
        st.header("เริ่มการจัดตารางสอน")
        st.markdown("โปรดใส่ข้อมูลด้านล่างเพื่อใช้ในการสร้างตารางเรียนตารางสอน")

    #Adjusting Conditions

        # number of session perday
        num_sessions_per_day = st.number_input('จำนวนคาบต่อวัน (ไม่นับคาบลดเวลาเรียนเพิ่มเวลารู้)',min_value = 4, max_value = 8, value = 5, key='sessionsaday')
        # Homeroom Time
        # วันที่ คาบที่
        Homeroom_day = st.selectbox('วันที่มีคาบแนะแนว (Homeroom) คือวัน',('จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์'), key = 'homeroomday')
        st.markdown("โปรแกรมจะจัดให้คาบแนะแนวเป็นคาบแรกของ**วัน{}**".format(st.session_state.homeroomday))
        
        #boyscout day
        boyscout_day = st.selectbox('วันที่มีคาบลูกเสือคือวัน',('จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์'), index = 3, key = 'boyscoutday')     

        # PE day
        PE_day = st.multiselect('วันที่มีคาบวิชาพละคือวัน (เลือกได้มากกว่า 1 วัน)', ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์'],['อังคาร'], key='PEday')
        if PE_day == []:
            st.error("โปรดเลือกวันที่มีคาบวิชาพละ")

        if boyscout_day in PE_day:
            st.error("วันสอนลูกเสือและพละเป็นวันเดียวกัน โปรดเลือกวันใหม่")
        
        st.write('**ตัวเลือกขั้นสูง**')
        
        #Morning and Afternoon Class preference

        with st.expander('วิชาที่อยากให้สอนช่วงเช้าและช่วงบ่าย'):
            morning_sessions_per_day = st.number_input('จำนวนคาบตอนเช้า',min_value = 2, max_value = 5, value = int(np.ceil(st.session_state.sessionsaday/2)), key='morningsessionsaday')
            morning_class = st.multiselect('วิชาที่อยากให้เรียนตอนเช้า (เรียงตามความสำคัญ)', pd.Series(Subjects),['คณิตศาสตร์','ภาษาไทย','ภาษาอังกฤษ'], key='morningclass')
            afternoon_class = st.multiselect('วิชาที่อยากให้เรียนตอนบ่าย (เรียงตามความสำคัญ)', pd.Series(Subjects),['พลศึกษา','ศิลปะ','การงานอาชีพ','สุขศึกษา'], key='afternoonclass')
            
        # people need to cover the class ex. ill, pragnancy, or retire
        # 2 modes: cover by one or more people/ homeroom teacher takecare of students

        with st.expander('แผนสำรองกรณีคุณครูมีความเสี่ยงที่จะไม่สามารถสอนได้'):
            st.write('ในบางครั้งครูบางท่านมีความเสี่ยงสูงที่จะไม่สามารถสอนได้ โดยสาเหตุอาจมาจาก อาการเจ็บป่วยเรื้อรัง การตั้งครรภ์ การเกษียณ หรือเนื่องจากจำนวนครูไม่เพียงพอ ท่านสามารถออกแบบแผนสำรอง โดยให้ครูท่านอื่นช่วยคุมชั้นเรียนแทน')
            risky_absent_teacher = st.multiselect('คุณครูที่มีความเสี่ยงที่จะไม่สามารถสอนได้', pd.Series(Teacher), key='absentteacher')

            risk_management = {t:'' for t in st.session_state.absentteacher}
            for t in st.session_state.absentteacher:
                st.write('แผนสำรองสำหรับครู{}'.format(t))
                options1 = ['ครูประจำชั้น','กำหนดเอง']
                options2 = [k for k in Teacher if k != t]
                if TA[t]['ประจำชั้น'] !='':
                    options1 = ['กำหนดเอง']
                bplan = st.selectbox('ครูที่จะดูแลห้องเรียนให้แทน', pd.Series(options1), key ='Bplanfor{}'.format(t))
                if bplan == 'กำหนดเอง':
                    risk_management[t] = st.multiselect('ครูที่จะดูแลห้องให้แทน', pd.Series(options2), key = 'groupBplanfor{}'.format(t))
                    st.write('โปรแกรมจะจัดคาบสอนของครู{} ให้ไม่ตรงกับครูที่เลือก'.format(t))
                else:
                    risk_management[t] = 'ครูประจำชั้น'
                    st.write('โปรแกรมจะจัดคาบสอนของครู{} ให้ไม่ตรงกับครูประจำชั้นของนักเรียนที่เรียนคาบนั้น ๆ'.format(t))

            teacher_planB_homeroom = [t for t in st.session_state.absentteacher if risk_management[t] == 'ครูประจำชั้น']
            teacher_planB_group = [t for t in st.session_state.absentteacher if risk_management[t] != 'ครูประจำชั้น']

    #Scheduling
        Days = ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์']
        Timeslot = [(i,j) for i in Days for j in range(st.session_state.sessionsaday)]

    # ''' Model formulation'''

        p = pu.LpProblem('School Timetabling Problem', pu.LpMinimize)

        # Add variables
    
        var = pu.LpVariable.dicts("ClassAtTime", (Classes,Timeslot), lowBound = 0, upBound = None, cat='Binary')
        # st.write(var.name)


        # '''Soft Constraints'''  
        Penelty_weight = 100
        Reward = -10

        # Morning Class Preference
        Constraint_var = []
        for cc in st.session_state.morningclass:
            Dummy = []
            for j in range(st.session_state.sessionsaday):
                Dummy.append(pu.lpSum(var[c][(i,j)] for i in Days for c in subject_plan[cc]))
            Constraint_var.append(Dummy)

        Penelty_distribution = []
        for c in st.session_state.morningclass:
            Dummy = []
            for i in range(st.session_state.sessionsaday):
                if i+1 <= st.session_state.morningsessionsaday:
                    Dummy.append(1/np.power(2,i+st.session_state.morningclass.index(c))*Reward)
                else:
                    Dummy.append(1/np.power(2,st.session_state.morningclass.index(c))*np.power(Penelty_weight,i))
            Penelty_distribution.append(Dummy)
            
        Morning_Class_Penelty = pu.lpSum(np.dot(Penelty_distribution[i],Constraint_var[i]) for i in range(len(st.session_state.morningclass)))

        # Afternoon Class Preference

        Constraint_var = []
        for cc in st.session_state.afternoonclass:
            Dummy = []
            for j in range(st.session_state.sessionsaday):
                Dummy.append(pu.lpSum(var[c][(i,j)] for i in Days for c in subject_plan[cc]))
            Constraint_var.append(Dummy)

        Penelty_distribution = []
        for c in st.session_state.afternoonclass:
            Dummy = []
            for i in range(st.session_state.sessionsaday):
                if i+1 <= st.session_state.sessionsaday-st.session_state.morningsessionsaday:
                    Dummy.insert(0,1/np.power(2,i+st.session_state.afternoonclass.index(c))*Reward)
                else:
                    Dummy.insert(0,1/np.power(2,st.session_state.afternoonclass.index(c))*np.power(Penelty_weight,i))
            Penelty_distribution.append(Dummy)
            
        Afternoon_Class_Penelty = pu.lpSum(np.dot(Penelty_distribution[i],Constraint_var[i]) for i in range(len(st.session_state.afternoonclass)))
                
        # Add Objective
        p += (Morning_Class_Penelty+Afternoon_Class_Penelty,"Sum_of_Total_Penalty",)

        # '''Hard Constraints'''

        # Teaching According to the Curriculum
        for c in Classes:
            p += (pu.lpSum(var[c][s] for s in Timeslot) == 1)

        # Teachers teach one class at a time and Plan B ครูประจำชั้น สำหรับครูเสี่ยงลาบ่อย
        # When kindergarten teacher teach English, Home room teacher need to take care kindergarten instead.
        
        for t in Teacher:
            t_class = []
            for c in Classes:
                if t == c[0] or (TA[t]['ประจำชั้น'] == c[1] and c[0] in teacher_planB_homeroom):
                    t_class.append(c)
            for s in Timeslot:
                p += (pu.lpSum(var[c][s] for c in t_class) <= 1)
            if t in teacher_planB_group:
                care_class =[]
                for c in Classes:
                    if c[0] in risk_management[t]:
                        care_class.append(c)
                for s in Timeslot:
                    p += (pu.lpSum(var[c][s] for c in care_class) + pu.lpSum(var[c][s] for c in t_class) <= len(risk_management[t]))


        # Each Grade attend one class at a time
        for g in Grades:
            for s in Timeslot:
                p += (pu.lpSum(var[c][s] for c in student_plan[g]) <= 1)

        # Home room class during the first period of assigned morning
        for c in Classes:
            p += (var[c][Timeslot[Days.index(st.session_state.homeroomday)*st.session_state.sessionsaday]] == 0)

        # Not Learn the same subject on the same day_Hard Constraint Version
        for s in Subjects:
            for g in Grades:
                sg_class = []
                for c in Classes:
                    if s == c[2] and g ==c[1]:
                        sg_class.append(c)
                for i in Days:
                    p += (pu.lpSum(var[c][(i,j)] for c in sg_class for j in range(st.session_state.sessionsaday)) <= 1)

        # Clases distribute close to average everyday for teacher
        sensitivity = 1
        for t in Teacher:
            t_average_per_day = len(teaching_workload[t])/len(Days)
            for i in Days:
                p += (pu.lpSum(var[c][(i,j)] for c in teaching_workload[t] for j in range(st.session_state.sessionsaday)) <= t_average_per_day + sensitivity)

        # PE Day
        PE_Not_Teach_Day =[d for d in Days if d not in st.session_state.PEday]

        for i in PE_Not_Teach_Day:
            p += (pu.lpSum(var[c][(i,j)] for c in subject_plan['พลศึกษา'] for j in range(st.session_state.sessionsaday)) == 0)

        # '''Solve'''

        # Save and Solving

        st.write('เมื่อเตรียมข้อมูลพร้อมแล้ว คุณครูสามารถกดปุ่มด้านล่างเพื่อเริ่มสร้างตารางสอนได้เลย :sunglasses:')

        solve_button = st.button('เริ่มการสร้างตารางสอน')

        if solve_button:
            with st.spinner('...โปรดรอ...'):
                p.solve()

                if pu.LpStatus[p.status] == 'Infeasible':
                    st.error("โปรแกรมไม่สามารถหาคำตอบที่สอดคล้องกับทุกเงื่อนไขได้ โปรดตรวจสอบข้อมูลที่ให้ หรืออาจปรับเงื่อนไขแผนสำรองโดยลดจำนวนครูที่เสี่ยงลาหรือเพิ่มจำนวนครูที่จะช่วยดูแลห้องแทน")
                else:
                    colums_name = ['คาบที่ {}'.format(i+1) for i in range(st.session_state.sessionsaday)]
                    Last_period = []
                    for i in range(len(Days)):
                        if Days[i] != st.session_state.boyscoutday:
                            Last_period.append('ลดเวลาเรียน ฯ')
                        else:
                            Last_period.append('ลูกเสือ')

            
                    df_teacher = { t : pd.DataFrame(columns= colums_name) for t in Teacher}   

                    for t in Teacher:
                        for i in Days:
                            Dplan = []
                            for j in range(st.session_state.sessionsaday):
                                Dummy = [round(var[c][(i,j)].varValue) for c in teaching_workload[t]] #somehow, some solution is not exactly one.
                                # print(sum(Dummy))
                                if sum(Dummy) == 0:
                                    Dplan.append('')
                                else:
                                    for c in teaching_workload[t]:
                                        if round(var[c][(i,j)].varValue) == 1:
                                            Dplan.append(c[2]+' '+c[1])
                            # print(Dplan)
                            # print(len(Dplan))
                            df_teacher[t].loc[i] = Dplan

                        for g in Grades: #homeroom teacher teach homeroom at the first period of assigned day
                            if g == TA[t]['ประจำชั้น']:       
                                df_teacher[t].at[st.session_state.homeroomday,colums_name[0]] = 'แนะแนว {}'.format(g)
                        
                        # ['ลดเวลาเรียน ฯ (ต้านทุจริตศึกษา)','ลดเวลาเรียน ฯ (ม่วนซื่นโฮแซว)','ลูกเสือ','ลดเวลาเรียน ฯ (ตามรอยเถ้าแก่น้อย)','ลดเวลาเรียน ฯ (คนดีมีคุณธรรม)']

                        df_teacher[t].insert(3,'พักรับประทานอาหารกลางวัน',['','','','','']) # Insert Lunch Time
                        df_teacher[t].insert(st.session_state.sessionsaday+1,'คาบที่ {}'.format(st.session_state.sessionsaday+1),Last_period) # Insert less study more knowledge
                        
                    # Student Schedule Dataframe
                    df_student = { g : pd.DataFrame(columns=colums_name) for g in Grades}   

                    for g in Grades:
                        for i in Days:
                            Dplan = []
                            for j in range(st.session_state.sessionsaday):
                                Dummy = [round(var[c][(i,j)].varValue) for c in student_plan[g]] #somehow, some solution is not exactly one.
                                if sum(Dummy) == 0:
                                    Dplan.append('')
                                else:
                                    for c in student_plan[g]:
                                        if round(var[c][(i,j)].varValue) == 1:
                                            Dplan.append(c[2])
                            df_student[g].loc[i] = Dplan

                        df_student[g].at[st.session_state.homeroomday,colums_name[0]] = 'แนะแนว'
                        df_student[g].insert(3,'พักรับประทานอาหารกลางวัน',['','','','','']) # Insert Lunch Time
                        df_student[g].insert(st.session_state.sessionsaday+1,'คาบที่ {}'.format(st.session_state.sessionsaday+1),Last_period) # Insert less study more knowledge

                    for g in Grades[:3]: 
                        df_student[g].replace('พลศึกษา','สุขศึกษา/พลศึกษา',inplace = True)


                    with st.expander("ตรวจดูตารางเรียนและตารางสอน"):
                        for g in Grades:
                            st.write('นักเรียนชั้น {}'.format(g))
                            st.write(df_student[g])
                        for t in Teacher:
                            st.write('ครู{}'.format(t))
                            st.write(df_teacher[t])


                # '''Create a Pandas Excel writer using XlsxWriter as the engine.'''
                    
                    writer = pd.ExcelWriter('ตารางสอน{} เทอม {} ปีการศึกษา {}.xlsx'.format(school_name,semester,year), engine='xlsxwriter')

                    # Write each dataframe to a different worksheet.
                    for t in Teacher:
                        df_teacher[t].to_excel(writer, sheet_name='ครู{}'.format(t))
                    for g in Grades:
                        df_student[g].to_excel(writer, sheet_name='นักเรียนชั้น {}'.format(g))

                    # Close the Pandas Excel writer and output the Excel file.
                    writer.save()

                    with open('ตารางสอน{} เทอม {} ปีการศึกษา {}.xlsx'.format(school_name,semester,year), 'rb') as my_file:
                        st.download_button(
                            label = '📥 ดาวน์โหลดตารางสอน 📥', 
                            data = my_file, 
                            file_name = 'ตารางสอน{} เทอม {} ปีการศึกษา {}.xlsx'.format(school_name,semester,year), 
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
