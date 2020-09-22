Thanks for like downloading this, and stuff...
	
	 	     -------------------------------	
			MODEL GENERATOR VERSION 1.5
	  	     -------------------------------

This program allows you to create 3D models and look at them, and
stuff. The reason I wrote it is because I have made a few 3D games,
but I have nothing to make 3D models with, and making them in notepad
is really hard. So here is my model editor, which you can use if your
making games and the like. This read me is a kind of guide, to show you
how to do things.

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

				Simple Stuff
				------ -----

To make a box, make sure that the 'Add Box' button is selected, and
then drag on the main window (the white bit the the cross in it).
You'll see a doted line while you drag, and then a box when you
let go. You'll also see the boxes name appear in the list at the
top right. Change the writing in the text box below the list, and
press F12 or 'Set Name' to change the name of the box.

Each box in your model will have a name, and you select the name from
the list to edit that box. You'll have to give all the boxes decent
names so you can remember which is which later on.

Make another two boxes, so that you have three altogether. You'll
see that one box is darker, and has circles at each corner. This is
the selected box. Click the names if the other boxes in the list,
and the different boxes will become highlighted. Still with me?

There is a different method to do this - If you can see the box you
want to highlight, use the 'Select' tool and click on a corner of the
box. So long as another box dosn't have a corner in the same place, you
will highlight the box you want.

Now click on 'Move Box' and drag the mouse a few centimeters on the
main window. The highlighted box will move that distance across the
window. You don't have to have touched the shape. Just drag anywhere,
and the box will move that distance. Try it to see what I mean.

Now click on 'Move Vertex'. Move your mouse over the selected box, and
click on one of the corners. The circle will get darker. Now draw it
around the main window and release the mouse. The corner will have 
moved, changing the shape of the square! Wow! 

If you want to move one edge of your box around, you could either move
each vertex at a time, or you can use the 'Move Face' tool. Click on it,
and you'll see a circle showing you where the center of each face is on
the selected box. Click on a circle and draw it about, to move the whole
face. As with the 'Move Vertex' tool, holding shift while you drag will
only select one vertex or face. So if you have two faces exactly over
each other, while you would normally select both faces, holding shift
means you'll only select one.

See that tool bar thing at the top of the screen. Click on the
differnt buttons to change the view position. You can also use
F1 - F3, or select the view from the views menu. You probebly guesed,
but it lets you look at your model from the top, front or side. You
can edit your model in any mode in the same way.

If you want to spin a shape, use the spin function. If you click on
a corner of the selected shape, you'll spin it around that corner,
else it spins around the center of the box. Drag the mouse down
first, and you'll see a little box with a number in (Thats the angle
it'll spin throught) move the mouse around to change the angle to
spin it throught.

To remove a box, press delete, or chose delete from the edit menu.

To make your more complex models easier to understand, your can
change the colour of the boxes. This dosn't effect the compiled
model, just the model that you can edit. To change a bexes colour,
make sure it is highlighted, and right click on the main edit
window. A pop up menu wil appear. Choose the lower option,
'Change Colour', and select a colour from the diolog box. The box
that is highlighted, and all future boxes made will be that colour.

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

				Menus
				-----

The file options should be prety obvious. Merge loads a new model
on top of the previous one. Save, Save as, New, all obvoius.

You have got a copy and paste function in the edit menu, but it
dosn't use the windows clipboard, so if your running two copies
of MAKER at the same time, you can't copy between them. Sorry.

In the tools menu, Set name is just the F12 shortcut, to change the
name of boxes . Hide me minimised the window, and info counts the
number of boxes, corners and faces there are.

3D Viewer shows your model in 3D. You can't edit it, just move it
around the screen, by left draging, or spin it by right draging. 
Try it to get used to it. It only displays in wire frame, but its
only to give you an idea what it'll look like. There are two options
in the options menu that give you some choise as to how the model is
drawn.

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

			Changing Face properties
			-------- ---- ----------

Face properties is a complicated bit, so its got a whole section for it.

What you do, is go throught your model one box at a time, and select
which faces are hidden and which arn't.

If you want to edit a specific box, Right click in the main edit window, and
select the 'Edit surface properties' from the popup menu. Once in the porperties
window, you can press 'Skip to' to get a list of all the box names, and you can
skip to any one by clicking on that name

You can see which brush you are working on on the left, and each face 
in the box is numbered. On the right are two list boxes, each containing 
a tick box for each face. The top list box changes visibility, and the
bottom changes normals. Explinations of both follow --->

If you want a face to appear in the final model, make sure it has a tick
next to its number in the top list box. If the face is going to be hidden,
because another box will be totaly covering it, remove the tick from that
face, and it will be ignored when you compile the model. This is very handy
to reduce the size of your model, and increase the speed of which it can be
drawn.

If you now anything about 3D stuff, you should know about face normals, and
that you can remove all the faces that face away from you by checking the
face normal. The lower list box is for changing the face normals around.
Select the 'Remove hidden face' check box, and have a good look at your box
in the left window. If it looks wiard, try pressing 'Reverse visibility'
to swap the face normals around. You'll probebly use this button alot, cos
if you create a box by draging left to right, instead of right to left, the
box will end up inside out, and you'll need to reverse the normals to make
it look right. You can also reverse the normals of an individual face in a box
but I can't really see this being of any use to you. But you can still do it...


If you want a handy hint, make your model first, and look at it in 3D mode.
If parts don't look right, go throught the Remove Hidden Faces bit, and
swap the normals around on the boxes that don't look right. Once it looks
right, you can go around and remove hidden faces.

Hope that made sence to you all. If you find bugs, or are trying to do somthing
and can't let me know, so the next version can be better.

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

				Compiling	
				---------

Once you have made your model, and want to use it in a game, or somthing, 
you'll need to compile it. Select 'Export' from the file menu.
Click on 'Select File' to set a filename to export to.

There are two export options.

1) Simple mode. This stores each face with its own set of co-ordinate for  
   each vertex. This means that a cube, which has 6 faces, each with 4 corners,
   will have 24 points stored in it. That means that to rotate the cube, you'll
   need to rotate 24 points, not 8. Its the way most people learning how to do
   3D store shapes, before they learn better.

2) Complex mode. This picks out all the unique vertecies, stores them first, and
   then links each face corner to a one of the unique vertecies in the list. This
   means a cube will have 8 vertices stored, not 24 like in simple mode. Complex
   mode creates a file thats totally compatable with Direct3D. However, it takes
   alot longer that simple mode to compile, as it has to use tempory files, and
   all sorts of crap. It dosn't use arrays, so it can handle files of any size.

Once you set that lot up, press compile, to create your model. If you compile in
complex mode, it will ask you it you want to start the DirectX viewer. If you
click yes, you can use this program to look at your model in groud shading. You 
can also start this viewer, and the software texture mapper throught the 'Tools'/
'Open file viewer' menu. 
	

