# Design: the submission page

Derived from the built `site/index.html`. The world is Data Moments' own identity (taken from data-moments.com and the supplied logo), applied to an analyst's submission.

## Color
- Ink `#131313`: masthead, footer. Purple `#4A3F78`: headings, primary buttons, focused bars. Secondary text `#5E5486`.
- Cream `#FCFDF4` ground, `#F9FAE8` chart and note surfaces, white for the open accordion item.
- Gradient `135deg, #A2E5E7 to rgba(189,177,216,.89)`: executive-summary band, closing band, badges, the masthead GitHub pill. Never on text.
- Chart roles: lavender `#BDB1D8` (comparison), purple (focus store / Q4), teal `#2A9FA5` (rest of chain), teal `#A2E5E7` (third category). Link blue `#0A5BD8`.
- Hero field: deep violet-blue, radial violet and teal glows behind a cream card.

## Type
- Rubik 400/500/600 for everything (Hebrew and Latin). Headline 1.9 to 3.2rem, question titles 1.4rem, section sub-heads 1.12rem, body 17px (16px on phones), tables 15px. Numbers use tabular figures.

## Shape and depth
- 20px radius on accordion items, 18px on the hero card, 16px on chart surfaces, full pill on buttons. Hairline `rgba(74,63,120,.18)` borders; the open item gets a purple border and a soft offset shadow.

## Components
- Accordion item: gradient number badge, title, one-line answer always visible, chevron; opens with a grid-rows transition. Native `details` blocks hold "the calculation step by step".
- Charts are drawn in plain DOM: horizontal bars with zero line and reference line, stacked mix bars, dumbbell, paired bars; a segmented control switches the metric.

## Motion
- One entrance language: ease-out-expo (`cubic-bezier(.16,1,.3,1)`) for panel opening, bar growth and the badge turn. Disabled under `prefers-reduced-motion`.

## Accessibility
- Buttons with `aria-expanded`/`aria-controls`, regions labelled, focus ring `#0A5BD8`, selection and scrollbar themed, deep links (`#q1`), print opens every item and hides the step-by-step blocks.
