# Pivot Table Text Value Conversion
VBA macro for displaying text within Excel's pivot table values section 

<h3>Project Description</h3>
<p align="justify">Have you ever run into the pesky problem where any text-based data (that you add as a secondary column) in a pivot table incorrectly display as values instead of text? This is because pivot tables in Microsoft Excel are designed to only show numbers in the “Values” section. If you add a text field there, Excel displays a count of those text items by default. Meanwhile, it can only handle text fields in the “Row and Column” section (which is restricted only to the first column of the pivot table).
<br><br>
By running this VBA code as a macro/module in your workbook, it will effectively show text in the pivot table “Values” section by automating the following process:</p>
<ol>
  <li>Assign custom ID numbers to corresponding text fields/cells through two separate arrays. </li><br>
  <li>From there, these custom number formats are applied (via iteration loop) as conditional formatting rules where a numeric field (RegID) is added to the pivot table Values area, and summarized by the Max function. For example, this rule displays “Amphibian” if the cell contains the number 1.</li>
</ol>
<p align="justify">While the aforementioned technique can be applied manually, this multi-step task quickly becomes tedious and menial with larger data sets. One of the advantages of this macro is the efficient ease in which it converts the values to whatever corresponding text you define and assign.</p>
<h3>User Instructions for VBA Macro</h3>
<p align="justify">Before using this macro, make sure each category in the source data has a corresponding number and name. For example, let’s say the text-based data that you plan on displaying in the pivot table Values section is the type of animal associated to each listed species. Since there are 6 types of animals according to the scientific method of taxonomy classification, the following ID numbers were assigned to each type of animal in a separate column: 
<br>
<table>
  <tr>
    <th>Animal ID Number</th>
    <th>Animal Type</th>
  </tr>
  <tr>
    <td>1</td>
    <td>Amphibian</td>
  </tr>
  <tr>
    <td>2</td>
    <td>Mammal</td>
  </tr>
  <tr>
    <td>3</td>
    <td>Bird</td>
  </tr>
  <tr>
    <td>4</td>
    <td>Fish</td>
  </tr>
  <tr>
    <td>5</td>
    <td>Reptile</td>
  </tr>
  <tr>
    <td>6</td>
    <td>Invertebrate</td>
  </tr>
</table>
<br>
One quick way to achieve this is by applying a conditional equation to the numerical ID column that assigns these numbers based on the category (as demonstrated below). This numerical ID data is essentially the column that you will include in the pivot table in order to display as text.
<br><br>
The code has two sets of arrays that correspond to each other.  One currently contains 6 numbers to change to animal group categories (i.e. amphibian, animal, etc.). You can change those numbers and names, or add more to match your pivot table items.
<br><br>
For example, if you have 9 text-based categories, change the arrays to include all 9 items. Feel free to also rename the array name (i.e. AnimalGroup and AnimalIDNumber) to whatever you deem most appropriate.</p>

<pre>
  <code>
  'Define arrays lists
  AnimalGroup = _
    Array("Amphibian", "Mammal", "Fish", "Bird", "Reptile", "Invertebrate")
  AnimalIDNumber = Array(1, 2, 3, 4, 5, 6)
  </code>
</pre>
Update the array range in the iteration loop to match your desired items.
<pre>
  <code>
  'Iterate through array list to convert assigned values to associated text
  For i = 1 To 6
  </code>
</pre>

<h3>How to manually display text in pivot table values section:</h3>
<p align="justify">In case you’re interested in the full manual process (that this macro essentially replicates), I have also included the detailed directions below.</p>
<ol>
  <li>For the source text-based data that you want to display in the pivot table, make sure each text-based data category has a corresponding ID number. This ID number column essentially serves as a stand-in for the text data, and should be added to the Values section of the pivot table.</li><br>
  <li>When the ID number field is added to the Values section of the pivot table, Excel automatically sets the summary function to Sum. Instead of a sum of the ID numbers, we want to see the actual ID numbers by:</li>
    <ul>
      <li>Right-clicking on one of the value cells.</li>
      <li>In the pop-up menu, click “Summarize Values By”, and then select “Max”. The pivot table values will then change to show you correct ID number for each assigned text.</li>
    </ul><br>
  <li>To show these ID numbers as texts, you can combine these custom number formats with conditional formatting:</li>
    <ul>
      <li>Select all the value cells in the pivot table.</li>
      <li>On the Home tab ribbon in Excel, click “Conditional Formatting”.</li>
      <li>Click “New Rule” in order to open the New Formatting Rule dialogue box.</li>
      <li>In the “Apply Rule to” section, select the option that says “All cells showing Max of RegID values” for the categories/columns that you want to display as text. This option allows for flexible conditional formatting that will adjust if the pivot table layout changes.</li>
      <li>In the Select a Rule Type section, choose "Use a formula to determine which cells to format".</li>
      <li>In the formula box, enter or select the appropriate text-based cell that you assigned your 1st number ID and type the formula for Region ID: <code>=B2=1</code></li>
      <li>Click the Format button, then click the Number tab.</li>
      <li>In the Category list, click Custom.</li>
      <li>In the Type box, enter this custom number format: <code>[=1]"Amphibian";;</code></li>
        <ul>
          <li>The first part of the format tells Excel to show "Amphibian", for any positive numbers equal to 1.</li>
          <li>The 2 semi-colons are separators, and there is nothing in the 2nd section (negative numbers) or 3rd section (zeros) of the custom format.</li>
        </ul>
      <li>Click OK, twice, to close the windows.</li>
    </ul><br>
  <li>Repeat step 3 to create any additional conditional formatting rules for each remaining categories or text items. For example:</li><br>

<table>
  <tr>
    <th>(Animal) Number ID</th>
    <th>Conditional Formatting Formula</th>
    <th>Custom Number Format</th>
  </tr>
  <tr>
    <td>2</td>
    <td>=B2=2</td>
    <td>[=2]"Mammal";;</td>
  </tr>
    <tr>
    <td>3</td>
    <td>=B2=3</td>
    <td>[=3]"Bird";;</td>
  </tr>
    <tr>
    <td>4</td>
    <td>=B2=4</td>
    <td>[=4]"Fish";;</td>
  </tr>
    <tr>
    <td>5</td>
    <td>=B2=5</td>
    <td>[=5]"Reptile";;</td>
  </tr>
    <tr>
    <td>6</td>
    <td>=B2=6</td>
    <td>[=6]"Invertebrate";;</td>
  </tr>
</table>
