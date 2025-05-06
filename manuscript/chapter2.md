# Chapter 2: From Clicks to Code – Automating Excel Workflows

Excel is powerful, flexible, and familiar. But let’s be honest: most of what people do in Excel involves a dizzying blur of clicks, filters, copy-paste gymnastics, and manual chart setups. It works—until it doesn’t.

This chapter meets you right there: in the messy middle of spreadsheets. We’re not here to shame the manual steps. We’re here to **understand** them—then replace them with **reproducible, readable, and smarter automation** using PowerShell and a sprinkle of AI.

Because here’s the big shift:

> You don’t have to babysit Excel. You can script it.

And once you do? You start solving problems faster than you can click through them.

---

## 🧍 Manual Excel Steps: The Familiar Pain

Let’s start with what Excel users actually do, day in and day out:

* Import a CSV
* Clean up columns
* Filter out rows
* Create a Pivot Table
* Save it as a report

You’ve done this. Maybe hundreds of times. And if you haven’t automated it yet, it’s likely because the steps feel small. But they add up—and they’re fragile.

That’s where PowerShell enters.

---

## ⚡ The PowerShell Shortcut: `ImportExcel`

The [ImportExcel module](https://github.com/dfinke/ImportExcel) gives you a way to script all those manual steps—without opening Excel at all.

Here’s what that flow looks like in PowerShell:

```powershell
$outputReport = '.\Report.xlsx'
Remove-Item $outputReport -ErrorAction Ignore

$data = Import-Csv .\sales.csv

$data |
Where-Object { $_.Region -ne 'West' } |
Export-Excel $outputReport -WorksheetName Sales -AutoSize `
    -PivotRows Region -PivotData @{ Units = 'Sum' } -Activate -Show 
```

You just:

* Ingested a CSV dataset
* Filtered out rows
* Created a pivot table
* Saved a styled Excel report

And you did it in a few lines of code—**reproducibly**.

---

## 🧠 Sprinkle in AI: PSAISuite Joins the Workflow

Now imagine you want to describe the dataset or decide which pivot view makes the most sense. That’s not about syntax—it’s about **thinking**. This is where AI enters.

With [PSAISuite](https://github.com/dfinke/PSAISuite), you can pass your dataset to a model and ask:

```powershell
Get-Content .\sales.csv | Invoke-ChatCompletion -Messages "Summarize this dataset and suggest a pivot layout."
```

Or ask it to:

* Recommend filters
* Spot anomalies
* Propose calculated columns

This is light-touch AI. It doesn’t replace your logic. It **augments** it.

---



⚠️ When AI Isn’t Magic (Yet)

Models are helpful, not omniscient. And when they fumble, it’s usually for one of two reasons:

1. You’re throwing too much at it

LLMs don’t work like Excel—they don’t “see” your entire file. They read what you feed them, line by line. If your sheet is huge, filter it down first. Let the model focus.

2. The data's a mess

No headers? Inconsistent columns? Missing values? AI will try its best—but “best” might still be wrong. Models aren’t clairvoyant. They work with structure, not hope.

These aren’t dealbreakers. They're reminders: your prep shapes the AI's potential.

---

## 🏁 Wrap-Up: From Clicks to Confidence

You didn’t start this chapter with code—you started with the problem. The messy spreadsheet steps that are real and human and often frustrating.

By the end, you saw how to:

* Recognize repetitive steps
* Script them cleanly in PowerShell
* Use AI to generate insight or guidance

You’re no longer just using Excel. You’re shaping it into something smarter.

> And we’re just getting started.

---

Next up? We shift from **one-off scripts** to **agentic workflows**. Get ready to build Excel experiences that *think alongside you*.
