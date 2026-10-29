# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Static HTML/CSS/JS, one `site/index.html`, no build step (the user will deploy it on Vercel; `vercel.json` points the output directory at `site`). Chosen by the user's message: "deployment on Vercel". Fonts load from Google Fonts.

## Users

Primary: the hiring team at Data Moments (an Israeli data-analytics consultancy) who read the candidate's take-home submission and decide whether to interview. They read Hebrew, are analysts and managers, and want to see the five questions answered with numbers they can check. Secondary: the pet-store chain's management the analysis is written for.

## Product Purpose

The official submission of the "תרגיל חשיבה עסקית - אנליסט נתונים" take-home (candidate: Ofek Ben Porat). It must contain every requirement of the brief: answers to five questions, a data-quality description, assumptions, a short note on AI use and verification, and a way to follow the calculations (the GitHub repository). Success: a reviewer finds each answer within seconds, can open the evidence behind it, and can verify the numbers.

## Positioning

A submission that behaves like an analyst's deliverable: answer first, evidence on demand, every figure traceable to code in the repository.

## Operating Context

The brief allows up to three slides or a two-page document. This page replaces that document, so it stays short: one-line answers visible at all times, evidence expanded on click. Data are synthetic (2025, five stores, 41,685 receipt lines). Amounts in ILS before VAT.

## Capabilities and Constraints

Hebrew, right to left. Deployed on Vercel as static files. No external calls except web fonts. The repository is https://github.com/OfekBP/data-analyst-data-moments. All figures come from `dm_analyst/` and are checked against SQL and an Excel workbook by tests.

## Brand Commitments

The page is styled in Data Moments' own visual identity (the user's explicit request), taken from https://www.data-moments.com and the logo the user supplied (`site/assets/logo.png`): near-black top bar `#131313`, deep purple `#4A3F78` for headings and buttons, cream `#FCFDF4` / `#F9FAE8` surfaces, teal-to-lavender gradient (`#A2E5E7` to `#BDB1D8`, 135deg), link blue `#116DFF`, Rubik for headings, rounded pill buttons. The page is a candidate's submission, so it never speaks as Data Moments: it names the company only as the addressee.

## Evidence on Hand

`docs/answers.md` (the written answers), `outputs/` (tables), the repository, the Excel workbook. No testimonials, client logos or claims about Data Moments' services are used or invented.

## Product Principles

1. Answer first, evidence on demand.
2. Every number on the page can be traced to the repository.
3. Say what was tested, what was found, and what is only a hypothesis.
4. Short enough to read as a two-page document.

<!-- Facts above were inferred from the user's explicit brief and messages (no interview round answered); the stack and the candidate name were given by the user in chat. -->
