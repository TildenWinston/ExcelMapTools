Won first prize at R.U. Hacking 2021

## Inspiration
Recently I have been having to deal with map related data I found I was sorely missing a way to calculate distances and travel times between locations from within Excel. 

I have been searching for apartments and I wanted a function that could take the apartment address and calculate the commute time to work or the walking time to the nearest bus stop.

Similarly, I have been planning a long, multi-day trip and would like a function that would help me gauge if driving between two locations in a day is feasible.

## What it does
**mapDistance**

Input: Origin, Destination, API Key

mapDistance was the original function. It takes the two addresses and uses the Google directions API to get directions and other information. All that this function is interested in is the route distance. The route distance text version, which is already nicely formatted and includes units is returned as a string.

**mapDistanceRawVal**

Input: Origin, Destination, API Key

Same as mapDistance, but the distance is returned as an integer value in meters. This value is easier to manipulate programmatically.

**mapTime**

Input: Origin, Destination, API Key

Instead of returning distance, this function returns the time with units as a string.

**mapTimeRawVal**

Input: Origin, Destination, API Key

Same as mapTime, except time value is turned as an integer number of seconds. Once again, this is easier to manipulate than the string version. The value returned here will be the starting point for finding the closest (by time) destination out of a list of destinations.

**mapAllVal**

Input: Origin, Destination, API Key

This function returns all of the data in a single cell. While the data isn’t as pretty, parsing it directly in excel is easy enough and by returning multiple data points additional API calls can be avoided lowering costs. Trades off simplicity on the client side with cost on the server side.

**mapClosest**

Input: Origin, Array of addresses, API Key

This function returns time in minutes to the closest address in the array and it return the address in the format given of this closest point. If addresses are given in latitude and longitude, the same will be returned, etc. While the data isn’t as pretty, parsing it directly in excel is easy enough and by returning multiple data points additional API calls can be avoided lowering costs. Trades off simplicity on the client side with cost on the server side.
This function is slow and Excel will possibly hang. If Excel stops responding, give it a few minutes to complete the function.

**mapClosestMatrix**

Input: Origin, Array of addresses, API Key

This function returns time in minutes to the closest address in the array and it return the address in the format given of this closest point. If addresses are given in latitude and longitude, the same will be returned, etc. While the data isn’t as pretty, parsing it directly in excel is easy enough and by returning multiple data points additional API calls can be avoided lowering costs. Trades off simplicity on the client side with cost on the server side.
This function is much faster than mapClosest.
Limitations: A maximum of 25 addresses can be given in the array


**Notes on address format**
Addresses for origin or destination can be given as Lat Long Pairs, Location names (usually), or full addresses. Geocodes and other google maps formats should theoretically work, but these are not validated.

**Other notes**
Strings must be explicit with quotes

Transportation mode is an optional parameter, must be encased in quotes. 


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
* Closest destination function
* Easier deployment
* Switch to Azure to resolve GCP billing issues

Stretch functions:
* Different traffic models
* Different leave/arrival time options
* Elevation difference?
* Geocoding
* Place Details
* Time Zones?

## Try it for yourself!
1. Open Excel
2. Open VBA editor using alt + F11
3. Go to Tools > References and make sure "Microsoft Scripting Runtime", "Microsoft Internet Controls", and "Microsoft WinHTTP Services, version 5.1" are all selected and enabled
4. Add Module
5. Paste the code from [MappingFunctions.bas](https://raw.githubusercontent.com/TildenWinston/ExcelMapTools/main/MappingFunctions.bas) there

Alternatively, right click on the project tree in the left hand bar, click import file, and select MappingFunctions.bas

6. JSON.bas and jsonEXT.bas will both also needed to be added in this manner. These files provide the JSON parsing library.


Submitted to R.U. Hacking 2021 https://r-u-hacking-hackathon-2021.devpost.com/
https://devpost.com/software/excel-mapping-tool-kit


## Resources & References
* https://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
* https://stackoverflow.com/questions/15940047/passing-a-range-to-excel-user-defined-function-and-assigning-it-to-an-array
* 
