Employee_key = {
    # Dictionary of employee types and their corresponding short hand.
    'All Employees - All Employees' : 'All',
    'All Employees - FULL TIME EMPLOYEES': 'Full-Time',
    'All Employees - PART TIME EMPLOYEES': 'Part-Time',
    'Females - All Employees' : 'Female',
    'Females - FULL TIME EMPLOYEES' : 'Female Full-Time',
    'Females - PART TIME EMPLOYEES' : 'Female Part-Time',
    'Males - All Employees' : 'Male',
    'Males - FULL TIME EMPLOYEES' : 'Male Full-Time',
    'Males - PART TIME EMPLOYEES' : 'Male Part-Time'
    }

Published_tables_data = {
    # Dictionary of tables and there contributing data sources
    'Table test - Test data' : ['testtable'],
    'Table 1 - All Employees' : ['total'],
    'Table 2 - Occupation (2)' : ['occ1', 'occ2'],
    'Table 3 - Gor by Occ (2)' : ['gro', 'goc', 'wgor', 'occ1', 'occ2'],
    'Table 4 - Industry (2)' : ['iau', 'ibe', 'isc', 'igu', 'sc207', 'scd07'],
    'Table 5 - Gor by Ind' : ['gau','gbe', 'ggu', 'gsc', 'gi207', 'gri07', 'sc207', 'scd07', 'iau', 'ibe', 'isc', 'igu', 'wgor', 'wgb', 'weng', 'wew'],
    'Table 6 - Age' : ['agegroup', 'total', 'agenoadr'],
    'Table 7 - Work Geography' : ['warea', 'weng', 'wew', 'wgb', 'wgor', 'wlanew5'],
    'Table 8 - Home Geography' : ['harea', 'heng', 'hew', 'hgb', 'hgor', 'hlanew5'],
    'Table 9 - Work PC' : ['wpcnew5', 'weng', 'wew', 'wgor', 'wgb'],
    'Table 10 - Home PC' : ['hpcnew5', 'heng', 'hew', 'hgb', 'hgor'],
    'Table 11 - Place of work Travel To Work' : ['wttwnew5', 'weng', 'wew', 'wgb', 'wgor'],
    'Table 12 - Place of residence Travel To Work' : ['httwnew5', 'heng', 'hew', 'hgb', 'hgor'],
    'Table 13 - PubPriv' : ['ppr'],
    'Table 14 - Occ (4)' : ['occ1', 'occ2', 'occ3', 'occ4'],
    'Table 15 - Gor by Occ (3)' : ['go3'],
    'Table 15 - Gor by Occ (4)' : ['go4'],
    'Table 16 - Industry (4)' : ['iau', 'ibe', 'isc', 'igu', 'sc207', 'scd07', 'sc307', 'sc407'],
    'Table 20 - Age by Occ (2)' : ['agegroup', 'ag1', 'ag2'],
    'Table 21 - Age by Ind (2)' : ['agegroup', 'ai207', 'ad207', 'aau', 'abe', 'asc', 'agu'],
    'Table 25 - WGOR by PUB PRIV' : ['wgor', 'ppr', 'A85o'],
    'Table 26 - Care Workers (SOC 2000 code 6115)' : ['hc'],
    'Table 27 - WLEPS' : ['wl1', 'wl2'],
    'Table 28 - HLEPS' : ['hl1', 'hl2'],
    'Table 32 - WITL' : ['wi2', 'wi3', 'wgor'],
    'Table 33 - HITL' : ['hi2', 'hi3', 'hgor'],
    # NI Below
    'Table 3 - Gor by Occ (2) NI' : ['gro', 'goc', 'wgor'],
    'Table 5 - Gor by Ind NI' : ['gau', 'gbe', 'ggu', 'gsc', 'gi207', 'gri07', 'wgor'],
    'Table 6 - Age NI' : ['gag', 'wgor'],
    'Table 7 - Work Geography NI' : ['warea'],
    'Table 8 - Home Geography NI' : ['harea'],
    'Table 9 - Work PC NI' : ['wpcnew5'],
    'Table 10 - Home PC NI' : ['hpcnew5'],
    'Table 13 - PubPriv NI' : ['wgor', 'A85o'],
    'Table 15 - Gor by Occ (4) NI' : ['gro', 'goc', 'go3', 'go4', 'wgor']  
    }

Published_tables_templates = {
    # Tables and their respective template names
    'Table test - Test data' : 'Testtable template.xlsx',
    'Table 1 - All Employees' : 'Total template.xlsx',
    'Table 2 - Occupation (2)' : 'Occupation SOC20 (2) template.xlsx',
    'Table 3 - Gor by Occ (2)' : 'Work Region Occupation SOC20 (2) template.xlsx',
    'Table 4 - Industry (2)' : 'sic07 Industry (2) template.xlsx',
    'Table 5 - Gor by Ind' : 'sic07 Work Region Industry (2) template.xlsx',
    'Table 6 - Age' : 'Age Group template.xlsx',
    'Table 7 - Work Geography' : 'Work Geography template.xlsx',
    'Table 8 - Home Geography' : 'Home Geography template.xlsx',
    'Table 9 - Work PC' : 'Work Parliamentary Constituency template.xlsx',
    'Table 10 - Home PC' : 'Home Parliamentary Constituency template.xlsx',
    'Table 11 - Place of work Travel To Work' : 'Work Travel To Work Area template.xlsx',
    'Table 12 - Place of residence Travel To Work' : 'Home Travel To Work Area template.xlsx',
    'Table 13 - PubPriv' : 'Pubpriv template.xlsx',
    'Table 14 - Occ (4)' : 'Occupation SOC20 (4) template.xlsx',
    'Table 15 - Gor by Occ (3)' : 'Test Work Region Occupation SOC20 (3) template.xlsx',
    'Table 15 - Gor by Occ (4)' : 'Test Work Region Occupation SOC20 (4) template.xlsx',
    'Table 16 - Industry (4)' : 'sic07 Industry (4) template.xlsx',
    'Table 20 - Age by Occ (2)' : 'Age by Occupation SOC20 (2) template.xlsx',
    'Table 21 - Age by Ind (2)' : 'sic07 Age by Industry (2) template.xlsx',
    'Table 25 - WGOR by PUB PRIV' : 'Work Region PubPriv template.xlsx',
    'Table 26 - Care Workers (SOC 2000 code 6115)' : 'Care Workers SOC20 6135 & 6136 (Equiv. to SOC10 6145 & 6146, SOC 2000 6115) template.xlsx',
    'Table 27 - WLEPS' : 'Work LEPS template.xlsx',
    'Table 28 - HLEPS' : 'Home LEPS template.xlsx',
    'Table 32 - WITL' : 'Work ITL (3) template.xlsx',
    'Table 33 - HITL' : 'Home ITL (3) template.xlsx',
    # NI Below
    'Table 3 - Gor by Occ (2) NI' : 'Work Region Occupation SOC20 (2) NI template.xlsx',
    'Table 5 - Gor by Ind NI' : 'sic07 Work Region Industry (2) NI template.xlsx',
    'Table 6 - Age NI' : 'Work Region Age NI template.xlsx',
    'Table 7 - Work Geography NI' : 'NI WAREA template.xlsx',
    'Table 8 - Home Geography NI' : 'NI HAREA template.xlsx',
    'Table 9 - Work PC NI' : 'NI WPC template.xlsx',
    'Table 10 - Home PC NI' : 'NI HPC template.xlsx',
    'Table 13 - PubPriv NI' : 'Work Region PubPriv NI template.xlsx',
    'Table 15 - Gor by Occ (4) NI' : 'Work Region Occupation SOC20 (4) NI template.xlsx' 
    }

Published_table_breakdown = {
    #Variable type and their corresponding short hand
    #'Weekly pay - Gross' : 'GPAY',
    #'Weekly pay - Excluding overtime' : 'GPOX',
    'Basic Pay - Including other pay' : 'BPAYinc',
    #'Overtime pay' : 'OVPAY',
    #'Hourly Pay' : 'HE',
    #'Hourly pay - Excluding overtime' : 'HEXO',
    #'Annual pay - Gross' : 'AGP',
    #'Annual pay - Incentive' : 'ANIPAY',
    #'Paid hour worked - Total' : 'THRS',
    #'Paid hours worked - Basic' : 'BHR',
    #'Paid hours worked - Overtime' : 'OVHRS'
    }

Table_sub_number_key = {
    # Variable type and its corresponding sub table number
    'Weekly pay - Gross' : '.1',
    'Weekly pay - Excluding overtime' : '.2',
    'Basic Pay - Including other pay' : '.3',
    'Overtime pay' : '.4',
    'Hourly Pay' : '.5',
    'Hourly pay - Excluding overtime' : '.6',
    'Annual pay - Gross' : '.7',
    'Annual pay - Incentive' : '.8',
    'Paid hour worked - Total' : '.9',
    'Paid hours worked - Basic' : '.10',
    'Paid hours worked - Overtime' : '.11'
    }