# Idea Vault

I built Idea Vault because I wanted one place to capture writing and art ideas as they came to me. I send an idea to Telegram, and the workflow turns it into a structured entry in Google Sheets so I can find it again when I’m ready to work on it.

## How it works

1. A Telegram message lands in the **Raw Messages** tab.
2. Google Apps Script asks the OpenAI API to turn the message into a title, summary and category. It saves a master record in **All Tasks**.
3. The script files the idea under **Writing**, **Art**, **Professional Branding** or **Personal Branding**. When the category is unclear, it marks the entry **Needs Review**.

## A design decision

A Substack essay and an article pitch both belong in **Writing**, but each has a different next step. I used a separate destination field to capture that intent. An essay meant for Substack gets **Self Publish**, an article meant for a publication gets **Pitch**, and an idea without a clear publishing plan can stay **Undecided**.

[Walk through a sample Substack idea](examples/substack-idea.md) · [Read the Apps Script](idea-routing.gs)

This repository contains the categorisation and routing code. Telegram intake is handled separately.
