************* Nick's 3D Maze ************
****** Programmed by Nick Thompson ******
******-----------------------------******
******    Send any commentss to:   ******
******* NickaThompson@Hotmail.com *******
*******---------------------------*******

ADDING LEVELS
This game comes with a certain number of levels
and textures for those levels. While you can
create custom maps to load up, you can also add
to the level system.

Do do this name your level after the highest
level number in the maps directory. E.g. if the
highest is l21.map then name your level l22.map.

After doing this edit the file info.dat.  You
should see 3 numbers.  Increase the LAST number
by one and your level will be added onto the end
of the existing levels.

ADDING TEXTURES
This game is written so that the user can add
textures for the FLOOR and WALLS.

The way this is done is similar to adding levels.
Be careful when creating textures, view an existing
texture in the \pictures directory to see the
required dimensions.  You will notice there are two
copies, one with a white background and one with a
black background. Any area that is white on the
first AND black on the second is taken to be
transparent, that area will not be drawn on the
screen.

Name your texture one above the current highest.

Then edit the file info.dat and increase the first
number by one if you have created a floor tile.
Increase the second number by one if you have created
a wall.

Further help is given in the file info.dat


