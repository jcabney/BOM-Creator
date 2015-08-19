# BOM-Creator
Written in Excel Visual Basic 
Used to create product bill of material given product configuration code.
Initially, bill of material was created manually from parts lists in Excel. Program created to automate the process. 
Program efficiency not critical as the product volume is relatively low. 

The program has two parts. The first, Module1, takes assembly numbers and pulls in the part info for that assembly. 
The second, Module2, consolidates and sorts those parts into a final list.

The input required is the product code. The program goes to the “BOM Program” tab and looks at the particular assembly 
type, which tells it where to grab the parts from and what parts to grab. 

The assembly parts are added to the list below the product code. Once the parts for the assembly are printed to the
lines, the program goes down the list looking for assemblies. The parts for each assembly are printed and the 
process continues until all the assembly parts have been written.

After the initial list of parts has been created, the program scrolls through each assembly/part searching the 
list for identical parts, adding the quantities together, and deleting those parts. 
The assembly or part # with the total quantity is printed to the “Complied Part List” tab. This continues line by 
line, until the original list is blank. The part type is then added, and the list sorted.
