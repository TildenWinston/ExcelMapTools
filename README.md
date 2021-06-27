ExcelMapTools

## Inspiration
Recently I have been having to deal with map related data I found was sorely missing a way to calculate distances and travel times between locations from within Excel. 

I have been searching for apartments and I wanted a function that could take the apartment address and calculate the commute time to work or the walking time to the nearest bus stop.

Similarly, I have been planning a long, multi-day trip and would like a function that would help me gauge if driving between two locations in a day is feasible.

## What it does

## How we built it
The custom functions are written in Visual Basic for Applications. The functions call the Google Directions API which returns JSON to parse. VBA and Excel are not well equipped for dealing with JSON, so I used a third party library, [VBA-JSON parser](https://github.com/omegastripes/VBA-JSON-parser),  to parse it.

## Challenges we ran into
This project is my first attempt at using Visual Basic. I found the language was a bit tricky to get used to at first, but not too bad by the end. It was interesting because of the different syntax and because of how minimal the language is.

Parsing the JSON was harder than it should have been. It took me a while to get the JSON parsing property and to figure out how to use the new library to get data that goes below the first level.

## Accomplishments that we're proud of
5 basic, but working functions

## What we learned
I learned the basics of VBA and was reminded of how nice linting, autocomplete, and other modern IDE features are.

## What's next for Excel Mapping Tool Kit
Closest destination function

Submitted to R.U. Hacking 2021 - https://r-u-hacking-hackathon-2021.devpost.com/
