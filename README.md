In this automation we will deal with fixing a corrupted excel file.
Corrupted here mean:
1- Values of cells are spread on more than there are columns which means there are 6 columns but the values range is 9 columns or 10.
2- Even if the values are right under their column the values are mixed which means the date is under gender col and cost under department and so on.

So, what I did is:
1- shift the cells left under their columns.
2- use RegEx to match strings patterns and then replace values.

But HOW?
See the automation.
