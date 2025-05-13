# This example uses PSAISuite to analyze the sales data and suggest insights
# Make sure PSAISuite is installed: Install-Module PSAISuite

# Get the content of the CSV file and pass it to the AI for analysis
Get-Content .\sales.csv | Invoke-ChatCompletion -Messages "Summarize this dataset and suggest a pivot layout."

# Optional: You can also save the AI's suggestions to a file
# $aiSuggestions = Get-Content .\sales.csv | Invoke-ChatCompletion -Messages "Summarize this dataset and suggest a pivot layout."
# $aiSuggestions | Out-File .\ai_suggestions.txt
