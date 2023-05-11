import openpyxl

# Load the workbook
workbook = openpyxl.load_workbook('C:\cricket_scores.xlsx')

# Get the active sheet
worksheet = workbook.active

# Check if the "Points and NRR" sheet exists
if "Points and NRR" in workbook.sheetnames:
    points_worksheet = workbook["Points and NRR"]
else:
    # If the "Points and NRR" sheet doesn't exist, create a new sheet
    points_worksheet = workbook.create_sheet("Points and NRR")

# Define column headers for the second sheet
headers = ["Team Name", "Matches Played", "Win", "Loss", "Tie", "Points", "NRR"]

# Write headers to the second sheet if it's a newly created sheet
if points_worksheet.max_row == 1:
    points_worksheet.append(headers)

# Iterate over the existing sheet and update the values in the second sheet
for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, values_only=True):
    match_no = row[0]
    batting_1st_team = row[1]
    batting_2nd_team = row[2]
    batting_1st_team_total_runs = row[3]
    batting_1st_team_total_wickets_down = row[4]
    batting_1st_team_overs_played = row[5]
    batting_2nd_team_total_runs = row[6]
    batting_2nd_team_total_wickets_down = row[7]
    batting_2nd_team_overs_played = row[8]

    net_run_rate = 0.0
    if batting_1st_team_overs_played != 0 and batting_2nd_team_overs_played != 0:
        net_run_rate = (batting_1st_team_total_runs / batting_1st_team_overs_played) - (batting_2nd_team_total_runs / batting_2nd_team_overs_played)

    # Calculate win, loss, tie, and points based on the result
    if batting_1st_team_total_runs > batting_2nd_team_total_runs:
        win_1 = 1
        loss_1 = 0
        tie_1 = 0
        win_2 = 0
        loss_2 = 1
        tie_2 = 0
    elif batting_1st_team_total_runs < batting_2nd_team_total_runs:
        win_1 = 0
        loss_1 = 1
        tie_1 = 0
        win_2 = 1
        loss_2 = 0
        tie_2 = 0
    else:
        win_1 = 0
        loss_1 = 0
        tie_1 = 1
        win_2 = 0
        loss_2 = 0
        tie_2 = 1

    points_1 = win_1 * 2 + tie_1
    points_2 = win_2 * 2 + tie_2

  
    # Find the row index of the first team in the second sheet
    team_1_row_index = None
    for i, row in enumerate(points_worksheet.iter_rows(values_only=True)):
        if row[0] == batting_1st_team:
            team_1_row_index = i + 2  # Adding 2 to match the row index in the sheet (1 for header row)
            break

    # Update the corresponding row in the second sheet or append a new row if the team is not found
    if team_1_row_index is not None:
        matches_played_cell = points_worksheet.cell(row=team_1_row_index, column=2)
        if matches_played_cell.value is None:
            matches_played_cell.value = 0
        matches_played_cell.value += 1

        points_worksheet.cell(row=team_1_row_index, column=3).value = win_1
        points_worksheet.cell(row=team_1_row_index, column=4).value = loss_1
        points_worksheet.cell(row=team_1_row_index, column=5).value = tie_1
        points_worksheet.cell(row=team_1_row_index, column=6).value = points_1
        points_worksheet.cell(row=team_1_row_index, column=7).value = net_run_rate
    else:
        new_row = [batting_1st_team, 1, win_1, loss_1, tie_1, points_1, net_run_rate]
        points_worksheet.append(new_row)

    # Find the row index of the second team in the second sheet
    team_2_row_index = None
    for i, row in enumerate(points_worksheet.iter_rows(values_only=True)):
        if row[0] == batting_2nd_team:
            team_2_row_index = i + 2  # Adding 2 to match the row index in the sheet (1 for header row)
            break

    # Update the corresponding row in the second sheet or append a new row if the team is not found
    if team_2_row_index is not None:
        matches_played_cell = points_worksheet.cell(row=team_2_row_index, column=2)
        if matches_played_cell.value is None:
            matches_played_cell.value = 0
        matches_played_cell.value += 1

        points_worksheet.cell(row=team_2_row_index, column=3).value = win_2
        points_worksheet.cell(row=team_2_row_index, column=4).value = loss_2
        points_worksheet.cell(row=tea   m_2_row_index, column=5).value = tie_2
        points_worksheet.cell(row=team_2_row_index, column=6).value = points_2
        points_worksheet.cell(row=team_2_row_index, column=7).value = -net_run_rate  # Negative NRR for 2nd team
    else:
        new_row = [batting_2nd_team, 1, win_2, loss_2, tie_2, points_2, -net_run_rate]  # Negative NRR for 2nd team
        points_worksheet.append(new_row)

# Save the changes to the workbook
workbook.save('C:\cricket_scores.xlsx')