for sheet in wb:
    if sheet.name == 'Mapping':
        header_row_found = False
        tiuid_found = False  # TI Safety Mechanism Unique Identifier
        diagname_found = False  # Safety Mechanism Name
        category_found = False  # Safety Mechanism Category
        type_found = False  # Safety Mechanism Type
        interval_found = False  # Safety Mechanism Operation Interval
        testtime_found = False  # Test Execution Time
        action_found = False  # Action on Detected Fault
        report_found = False  # Time to Report
        # columns numbers below
        tiuid_col = 0  # TI Safety Mechanism Unique Identifier
        diagname_col = 0  # TI Diagnostic Name
        category_col = 0  # Safety Mechanism Category
        type_col = 0  # Safety Mechanism Type
        interval_col = 0  # Safety Mechanism Operation Interval
        testtime_col = 0  # Test Execution Time
        action_col = 0  # Action on Detected Fault
        report_col = 0  # Time to Report

        used_tiuid = set()  # hashset to check repeated tiuid
        for row in range(sheet.nrows):
            tiuid = ""
            for col in range(sheet.ncols):
                cval = ""
                try:
                    cval = re.sub('\n', ' ', sheet.cell(row, col))
                    cval = str(cval)
                except:
                    continue
