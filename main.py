import pandas as pd 


def main(data):

    assignments_df = pd.read_csv('data/assignments.csv').fillna('')

    ## keep MPs
    MPs = data['MPs']
    assignments_df = assignments_df[assignments_df['Term'].isin(MPs)]

    ## keep courses with grades
    assignments_df = assignments_df[assignments_df['Course']!='']
    

    ## drop assignments worth zero
    assignments_df = assignments_df[assignments_df['WorthPoints']!=0]

    ## drop assignments not graded yet
    assignments_df = assignments_df[assignments_df['RawScore'] != 'NG']
    assignments_df = assignments_df[assignments_df['RawScore'] != 'Ng']
    assignments_df = assignments_df[assignments_df['RawScore'] != 'ng']
    # drop excused assignments
    assignments_df = assignments_df[assignments_df['RawScore'] != 'EX']
    assignments_df = assignments_df[assignments_df['RawScore'] != 'Ex']
    assignments_df = assignments_df[assignments_df['RawScore'] != 'ex']
    assignments_df = assignments_df[assignments_df['RawScore'] != 'es']
    # drop assignments with no grade entered
    assignments_df = assignments_df[assignments_df['RawScore'] != '']
    # drop checkmarks
    assignments_df = assignments_df[assignments_df['RawScore'] != '✓']

    ## convert percentages
    assignments_df['Percent%'] = assignments_df.apply(convert_percentages, axis=1)
    

    ## adjusted worth points
    assignments_group_by_cols = ['Teacher', 'Assignment', 'Course', 'DueDate']
    assignments_dff = assignments_df.drop_duplicates(
        subset=assignments_group_by_cols+['Objective'])
    assignments_dff = assignments_dff.groupby(assignments_group_by_cols)[
        ['Objective', 'WorthPoints']].agg({'WorthPoints': ['min', 'max','sum'], 'Objective': 'nunique'}).reset_index()
    reassigned_cols = ['Teacher', 'Assignment', 'Course', 'DueDate',
                       'WorthPointsMax', 'WorthPointsMin', 'WorthPointsSum', 'ObjectivesCount']
    assignments_dff.columns = reassigned_cols

    assignments_df = assignments_df.merge(
        assignments_dff,
        on=assignments_group_by_cols,
        how='left'
    )

    assignments_df['WorthPoints'] = assignments_df.apply(
        recompute_worth_points, axis=1)

    assignments_df['numerator'] = assignments_df['WorthPoints'] * assignments_df['Percent%']
    assignments_df['denominator'] = assignments_df['WorthPoints']
    
    
    student_grades_df = assignments_df.groupby(
        ['StudentID', 'Course', 'Section','CategoryWeight'])[['numerator', 'denominator']].sum().reset_index()
    student_grades_df['Category%'] = student_grades_df['numerator'] / student_grades_df['denominator']
    student_grades_df['weighted%'] = student_grades_df['Category%'] * student_grades_df['CategoryWeight']

    student_grades_df = student_grades_df.groupby(['StudentID','Course','Section'])['weighted%'].sum().reset_index()


    student_grades_df['FinalMark'] = student_grades_df['weighted%'].apply(convert_final_mark)

    student_grades_df = student_grades_df.reset_index()
    
    student_grades_df['JupiterCourse'] = student_grades_df['Course']
    student_grades_df['JupiterSection'] = student_grades_df['Section']

    student_grades_df = student_grades_df[['StudentID', 'JupiterCourse','JupiterSection','FinalMark']]

    crossover_df = pd.read_csv('data/jupiter_crossover.csv')

    student_grades_df = student_grades_df.merge(
        crossover_df,
        on=['JupiterCourse','JupiterSection'],
        how='left',
    )

    

    ## open EGG

    egg_df = pd.read_excel('data/egg.xlsx').fillna('')
    
    egg_columns = egg_df.columns

    temp_egg_df = egg_df.copy()

    temp_egg_df = temp_egg_df.merge( 
        student_grades_df, 
        left_on=['StudentID','Course','Sec'],
        right_on=['StudentID','Course','Section'],
        how='left',
    )

    temp_egg_df['ReconciledMark'] = temp_egg_df.apply(
        reconcile_egg_and_jupiter, axis=1)
    

    writer = pd.ExcelWriter('output/egg_output.xlsx')
    temp_egg_df.to_excel(writer, sheet_name='EGG_Output')
    student_grades_df.to_excel(writer, sheet_name='JupiterComputedMarks')

    writer.close()

    return True

def recompute_worth_points(row):
    WorthPoints = row['WorthPoints']
    WorthPointsMax = row['WorthPointsMax']
    WorthPointsMin = row['WorthPointsMin']
    WorthPointsSum = row['WorthPointsSum']
    ObjectivesCount = row['ObjectivesCount']

    if ObjectivesCount == 0:
        return WorthPoints
    if WorthPointsMax == WorthPointsMin:
        return WorthPoints / ObjectivesCount
    else:
        return WorthPoints / WorthPointsSum

def reconcile_egg_and_jupiter(row):
    egg_mark = row['Mark']
    jupiter_mark = row['FinalMark']

    if egg_mark:
        return convert_final_mark(egg_mark)
    else:
        return jupiter_mark

def convert_final_mark(Mark):
    try:
        if Mark < 50:
            return 45
        if Mark < 65:
            return 55
        return round(Mark)
    except:
        return Mark

def convert_percentages(row):
    Percent = row['Percent']
    RawScore = row['RawScore']

    correction_dict = {
        '9!': 50,
        '3!!': 85,
        '1': 65,
        '41': 95,
        '21': 75,
    }

    if Percent in [100,95,85,75,65,50,45]:
        return Percent/100
    if Percent == 0:
        return 50/100
    try:
        if Percent < 100 and Percent > 45:
            return Percent/100
    except:
        Percent = correction_dict[RawScore]

    if Percent < 100 and Percent > 45:
        return Percent/100
    else:
        Percent = correction_dict[RawScore]
        return Percent/100

if __name__ == '__main__':
    data = {
        'MPs': ['S1-MP1',],
    }
    main(data)