import pandas as pd
import re

# To avoid warning of chained assignment
pd.options.mode.chained_assignment = None
# load Data from Excel Sheet (Given the file path as per your system)
df = pd.read_excel('C:/Users/Sanju/Downloads/Attendance.xlsx')

# Removing rows having all the NULL values
df1 = df.dropna(axis = 0,how='all')

# Removing columns having all the NULL values
df2 = df1.dropna(axis=1,how='all')

# Converting all the NaN values to 0 for better understanding
df3 = df2.fillna('0')

# Removing undesired rows and columns which are not required for final outcome
df4 = df3.drop(['Unnamed: 4','Unnamed: 6','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 16'], axis=1)

# To remove column containing Company
df11 = df4[df4['Unnamed: 1'] != 'Company:']
# To remove column containing Department
df12 = df11[df11['Unnamed: 1'] != 'Department:']
# Considering Total Duration for final output
df12["Indexes"]= df12["Unnamed: 1"].str.find('Total Duration', 0)


# Taking all the necessary data from df4 for desired output
a = df4[['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 7','Unnamed: 14','Unnamed: 15']][df4['Unnamed: 1'] == 'Emp Code:']
b = df4[['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 7','Unnamed: 14','Unnamed: 15']][df4['Unnamed: 14'] == 'Status']
c = df4[df4['Unnamed: 14'].isin(["Absent","Present "])]
d = df12[['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 7','Unnamed: 14','Unnamed: 15']][df12['Indexes'] == 0]
# Cleaned data for required output
df_row = pd.concat([a, b, c, d])

# This will sort a,b,c dataset 
df5 = df_row.sort_index()

# This will take only records containing Total Records and create new Data Frame
df_models_total_duration = pd.DataFrame(df5['Unnamed: 1'][3::4].tolist()).fillna('').add_prefix('total_duration_')

# hh will store value of Total Duration
hh = []
for k in range(0,len(df_models_total_duration['total_duration_0'])):
    hh.append(df_models_total_duration['total_duration_0'][k].split(','))

t1 = [''.join(tt) for tt in hh]

# Store values containing total duration
a = []
for u in range(len(t1)):
    a.append(t1[u].find('Total Duration', 0))

# Store values space separated 
m = []
for i in range(len(t1)):
    m.append(t1[i].split(' '))

# This will contains only Total duration and remove rest other information 
u = []
for i in range(0,len(m)):
    s = []
    for k in range(0,6):
        s.append(m[i][k])
    u.append(s)
    
# total_time will contains hour and minute for every employee
total_time = []
for i in range(len(u)):
    for j in range(len(u[i])):
        hrs = u[i][1].replace("Duration=","Duration = ")
        mts = u[i][3]    

    res = [(i) for i in hrs.split(' ') if i.isdigit()]
    total_time.append((res[0]))
    total_time.append(mts)

# fg contains time in standard formart as hh:mm
fg = []
for i in range(0,len(total_time),2):
    ttime = total_time[i]+':'+total_time[i+1]
    fg.append(ttime)

# max_time contains maximum time spend by employee in Door2 
max_time = max(fg)

# This will take only Punch Records and create new dataset
df_models = pd.DataFrame(df5['Unnamed: 15'][2::4].tolist()).fillna('').add_prefix('timing_')

# Convert data of df_models to a list
ff = []
for i in range(0,len(df_models['timing_0'])):
    ff.append(df_models['timing_0'][i].split(','))
    
# r1 is used for splitting of in and out data in final dataset
r1 = [''.join(ele) for ele in ff]

# This will make dataset of name of employees
df_models_name = pd.DataFrame(df5['Unnamed: 7'][0::4].tolist()).fillna('').add_prefix('name_')

# This will convert df_models_name to a list of name
gg = []
for i in range(0,len(df_models_name['name_0'])):
    gg.append(df_models_name['name_0'][i].split(','))
    
# n2 is used in final loop to concatenate name in the final dataset
n2 = [''.join(ele) for ele in gg]

# This will contains name of employee alongwith their total duration spend 
dict1 = {}

# This will loop till all the punch records are over can be applied on name as well
for i in range(len(r1)):
    
    # This will differentiate Door2 records from Door1 records
    time = [i.group() for i in re.finditer("\d{2}:\d{2}:(in|out)\(Door2\)",r1[i])]
    # "name" contains name along with Emp Code to be printed next to the Punch in and Punch out
    name = n2[i]
    l = []
    
    # This will store employee id & name and total duration spend in key value pair
    dict1[n2[i]] = fg[i]
    
    # To check there is an "out" for every "in" or not
    count = 2
    j = []
    # run for every punch time record and check for in and out 
    for i in time:
        # Convert uppercase 'in' to lowercase 'in'
        if i.lower().count('in'):
            # This will put timing of Door 2 in list j if contains 'in' else in 'out'
            if sum(map(lambda x:x.count('in'), j))==0:
                j.append(i)
                temp_k = i
            else:
                j.append('out')
                l.append(j)
                j = []
                j.append(i)
                count+=2
                continue
        # Convert uppercase 'Out' to lowercase 'out'
        if i.lower().count('out'):
            # Append Door 2 timing in list j if contains 'out' else in 'in'
            if temp_k.lower().count('in'): 
                j.append(i)
            else:
                j.extend(['in', i])
                l.append(j)
                j = []
                count+=2
                continue
            temp_k = i
        # Check 'in' for consecutive 'out' else put 'in' or 'out' accordingly
        if time[len(time)-1]==i and i.count('in'):
            j.append("out")
            l.append(j)
            continue
        if count%2!=0:
            l.append(j)
            j=[]
        count+=1

    final_data = pd.DataFrame(l, columns=['Punch - IN', 'Punch - OUT'])
    # Adding column 'Emp Code' in the Dataframe
    final_data['Emp Code'] = name
    final_data
    
    # Reframing the Data frame starting with Emp Code, Punch In and Punch out
    output = final_data.reindex(columns=['Emp Code','Punch - IN', 'Punch - OUT'])
    print(output)

# This will store only name of employee who spend maximum time using 
keys = [key for key, value in dict1.items()if value == max_time]

print("Maximum time spend by : {}, and total time is : {}. ".format(keys[0],max_time))
