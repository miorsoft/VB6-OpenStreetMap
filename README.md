
# OpenStreetMap

I dusted off an old program from more than a decade ago. (An that time the OSM data was really very poor...)

If I remember correctly it was published on PlanetSourceCode.
I revisited it and made a lot of improvements, from rendering (which uses RC6) to a general increase in performance.

It allows you to load and visualize an area/city/country with its streets and buildings using downloadable data from www.OpenStreetMap.com (saveing them to the \MAPS\ Folder with .osm extension).
Simply go to that site and export the data for the area of interest. (If you want to download large map click "Overpass API".)
The .osm file is interpreted and eleborated, then it is asked how many cars to place.

In fact, the beauty is that in addition to displaying the map, the movement of the cars (the traffic) is implemented. 
Each car starts from a random point and goes to a random point on the map. As soon as it arrives at its destination another random target is chosen.
This is effected by the shortest path search algorithm called Dijkstra.
It is possible to Right click on the map to set the target of the car that the "camera" is following.
Left click selects the nearest car, if "Follow Car" checked, otherwise it select the "camera lookat" position. 

You can Zoom in/out with MouseWheel.
The cars move around trying to avoid colliding (this still needs to be improved).
In case of collision, and during the movement of the cars a mini physics engine is implemented.
Rendering is very fast even for large maps.
~~The bottleneck is the number of the cars you put, especially if the map is large, as the Dijkstra algorithm is not as fast as a rocket.~~
When you load a large map you basically perform a kind of stress test of cairo graphics. :-)



If it was impossible to run you might not have the correct version of Microsoft XML (msxml6.dll)
However it should work even with earlier versions, perhaps replacing MSXML2.DOMDocument60 with MSXML2.DOMDocument or something


Some of the main points of interest:

- Dijkstra Algorithm (for finding the shortest paths between nodes in a weighted graph - https://en.wikipedia.org/wiki/Dijkstra's_algorithm)
- Conversion of Latitude Longitude to UTM (Universal Transverse Mercator coordinate system).
- Line Clipping (https://en.wikipedia.org/wiki/Line_clipping)
- Line Line intersection (the simplest way I found)



I think where possible I will make various improvements.

#### List of things TODO or things that someone could help me improve:
PERFORMANCE:
- Improve Dijkstra algorithm or implement A* algorithm suitable for connected nodes. (Something similiar to A* alredy done)
- Improve the spatialGird class or implement a QuadTree class or other kind of Spatial Hash algorithm.

RENDERING:
- Draw the road pavement differently depending on the type. (Now only asphalt/not asphalt).
- Others Ideas.

CARS:
- Implementing crossroads yeld/stop.
- Do a better driving collision avoidance system.


#### Other considerations:

Also then I would like to get suggestions on how to develop the program further, possible directions to take, such as:
- If going in a game direction expose any rule suggestions.
- 3D rendering.
- Other.

Anyway any sort of suggestions and contributions or bug detect/fix will be greatly appreciated.

Have Fun!

PS: Delte ZipExtracted.txt file or change its text from 1 to 0 to perform zip extraction at First time Run
