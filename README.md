# ExcelOpinionMining

This is a Java file that I created when I was helping a friend do some opinion mining with an Excel file. 

It's very simple and stupid, yet the complexity is very high, because we parse the Excel file, then we are searching in the new Map structure
if that tweet, or some similar tweet, it's in the structure. While we are doing that we add to another Map structure the words in the tweets
and later we sort it to see the most used words.

At the end, it prints everything in a txt file.

# Dependencies

Apache POI
Apache Commons Lang
