# Inbox Assistant

Turns messy customer messages into a clean formatted spreadsheet.

![Screenshot of the output spreadsheet](inbox-assistant-results-spreadsheet.png)

## The problem

Small businesses get enquiries as unstructured text: emails, texts, web forms. Someone has to read each one, work out who it's from, what they want and how urgent it is, and then type it up somewhere themselves, this program automates this.

## What it does

1. Reads customer messages from a text file, one message per line
2. Sends each one to Claude (Haiku) with instructions to extract structured data: type, name, company, phone, email, request, deadline, contact notes.
3. Uses Python to work out urgency from the deadline.
4. Writes everything to a formatted Excel file

## How to run it

1. Install Python 3.10+ and the dependencies:
   '''
   pip install -r requirements.txt
   '''

2. Get an API key from console.anthropic.com and set it as an environment variable (never put it in the code):
   '''
   setx ANTHROPIC_API_KEY "your-key-here"
   '''
   Then open a new terminal.

3. Put fake test messages in `messages.txt`, one per line.

4. Run:
   '''
   python inbox_assistant.py
   '''

5. A timestamped `.xlsx` file appears in the folder.

## How I tested it

I wrote 30 deliberately awkward messages covering: spam and scams, a Spanish message, prompt injection, self-corrections, two people in one message, spaced-out phone numbers, multiple phone numbers, vague and relative dates, and emoji-only messages. After each prompt change I re-ran all 30 to check nothing that had been correct got worse.

## Design Decisions

- AI handles what its good at (the language), plain code handles the date logic
- Failed messages show up in the spreadsheet as "NEEDS REVIEW" instead of being lost.

## What I would do next

- Connect a real source to use real data
- Add a review step where a human can correct AI output


## Known limitations

- **One deadline per message.** A message with three items and three deadlines gets squashed into one.
- **Ambiguous dates** like "next Friday" are interpreted one way with no way to ask for clarification.
- **Results aren't perfectly repeatable.** The model can classify borderline messages differently between runs (e.g. emoji-only messages).
- **Input is a text file only.** There's no live email or web form connection yet.
- Tested on fake data only.

## Tech

Python, Anthropic API (Claude Haiku), openpyxl.
