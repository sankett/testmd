---


---

<h1 id="existing-vb.net-vsto-form-analysis">Existing <a href="http://VB.NET">VB.NET</a> VSTO Form Analysis</h1>
<h2 id="form-functionality">Form Functionality</h2>
<p>The form (<code>frm_CoverDateChange</code>) provides a simple user interface to update the date on a PowerPoint cover page. Its main functionalities include:</p>
<ul>
<li>
<p><strong>Date Selection:</strong><br>
A  <code>DateTimePicker1</code>  control lets users pick a date.</p>
</li>
<li>
<p><strong>Date Format Selection:</strong><br>
Four radio buttons allow the user to choose a display format for the date:</p>
<ul>
<li>“Month Year” (<code>optMMMYYYY</code>)</li>
<li>“Month Day, Year” (<code>optMMMDDYYYY</code>)</li>
<li>“Day Month, Year (EU)” (<code>optDDMMMYYYY_EU</code>)</li>
<li>“No Date” (<code>optNoDate</code>)</li>
</ul>
</li>
<li>
<p><strong>Dynamic Date Format Update:</strong><br>
The  <code>DateTimePicker1</code>  control’s custom format is updated dynamically when the user changes the radio button selection. The format changes based on the radio button checked state:</p>
<ul>
<li>If  <code>optMMMYYYY</code>  is checked → Format becomes  <code>"MMMM yyyy"</code>.</li>
<li>If  <code>optMMMDDYYYY</code>  is checked → Format becomes  <code>"MMMM dd, yyyy"</code>.</li>
<li>If  <code>optDDMMMYYYY_EU</code>  is checked → Format becomes  <code>"dd MMMM, yyyy"</code>.</li>
<li>If  <code>optNoDate</code>  is checked → Format is set to a blank string.</li>
</ul>
</li>
<li>
<p><strong>Regional Defaults:</strong><br>
On load, the form checks a template property (using  <code>getTemplateProperties</code>) to determine the region. If the region is Europe (<code>REGION_EUROPE</code>), the EU date format is chosen; otherwise, the default “Month Year” format is set.</p>
</li>
<li>
<p><strong>PowerPoint Integration:</strong><br>
On clicking the OK button, the form calls  <code>Cover.setCoverPageDate</code>  with the string from  <code>DateTimePicker1</code>, updating the active PowerPoint presentation’s cover page.</p>
</li>
</ul>
<hr>
<h2 id="methods-events-and-properties">Methods, Events, and Properties</h2>
<h3 id="methods">Methods</h3>
<ul>
<li>
<p><strong>frm_CoverDateChange_Load(sender, e):</strong></p>
<ul>
<li>Triggered when the form loads.</li>
<li>Checks the region via  <code>getTemplateProperties(TemplateProperty.REGION)</code>  to set an initial date format.</li>
<li>Updates the  <code>DateTimePicker1</code>  and selects the appropriate radio button.</li>
</ul>
</li>
<li>
<p><strong>btnOk_Click(sender, e):</strong></p>
<ul>
<li>Invoked when the “Ok” button is clicked.</li>
<li>Calls  <code>Cover.setCoverPageDate</code>  passing the selected date (as text).</li>
</ul>
</li>
<li>
<p><strong>optXXX_CheckedChanged(sender, e):</strong></p>
<ul>
<li>A shared handler for the CheckedChanged events of all four radio buttons.</li>
<li>Updates the  <code>DateTimePicker1.CustomFormat</code>  property depending on which radio button is checked.</li>
</ul>
</li>
</ul>
<h3 id="events">Events</h3>
<ul>
<li>
<p><strong>Form Load Event:</strong><br>
Handled by  <code>frm_CoverDateChange_Load</code>.</p>
</li>
<li>
<p><strong>Button Click Event (OK):</strong><br>
Handled by  <code>btnOk_Click</code>.</p>
</li>
<li>
<p><strong>Radio Button CheckedChanged Events:</strong><br>
Handled for each radio button (<code>optMMMYYYY</code>,  <code>optMMMDDYYYY</code>,  <code>optDDMMMYYYY_EU</code>,  <code>optNoDate</code>) by the same method to update the date format.</p>
</li>
</ul>
<h3 id="properties--controls">Properties / Controls</h3>
<ul>
<li>
<p><strong>Controls on the Form:</strong></p>
<ul>
<li><code>DateTimePicker1</code>: Date picker for choosing dates.</li>
<li><code>btnOk</code>: Button to apply the selected date.</li>
<li><code>btnClose</code>: Button to cancel or close the form.</li>
<li><code>PictureBox1</code>: Displays an image (<code>Date_1</code>) related to date selection.</li>
<li><code>GroupBox1</code>: Contains the date-related controls.</li>
<li><code>optMMMYYYY</code>,  <code>optMMMDDYYYY</code>,  <code>optDDMMMYYYY_EU</code>,  <code>optNoDate</code>: Radio buttons to choose different date display formats.</li>
</ul>
</li>
<li>
<p><strong>Active Presentation (<code>presActive</code>):</strong><br>
Private variable holding the active PowerPoint presentation from  <code>Globals.ThisAddIn.Application.ActivePresentation</code>.</p>
</li>
</ul>
<hr>
<h2 id="connections--dependencies">Connections / Dependencies</h2>
<ul>
<li>
<p><strong>VSTO (Visual Studio Tools for Office):</strong><br>
Uses the PowerPoint interop via  <code>Microsoft.Office.Interop.PowerPoint</code>  to interact with the active presentation.</p>
</li>
<li>
<p><strong>Globals.ThisAddIn:</strong><br>
Accesses the global add-in instance to obtain the active presentation.</p>
</li>
<li>
<p><strong>Cover Module/Class:</strong><br>
The method  <code>Cover.setCoverPageDate</code>  is used to update the cover page with the selected date. This module acts as the connector between the UI and the PowerPoint presentation.</p>
</li>
<li>
<p><strong>getTemplateProperties Function:</strong><br>
Retrieves template settings (e.g., regional settings) which influence the form’s behavior at load time.</p>
</li>
<li>
<p><strong>Custom Controls (myButton, myRadioButton):</strong><br>
Custom implementations of buttons and radio buttons used for consistent UI styling.</p>
</li>
</ul>
<hr>
<h1 id="implementation-plan-for-migrating-to-office.js">Implementation Plan for Migrating to Office.js</h1>
<p>Migrating from a VSTO-based form to an Office.js add-in involves rethinking the UI in web technologies (HTML, CSS, JavaScript) and replacing interop calls with Office.js APIs.</p>
<h2 id="migration-steps">Migration Steps</h2>
<ol>
<li>
<p><strong>Project Setup</strong></p>
<ul>
<li>Create a new Office.js PowerPoint add-in project using the Yeoman generator or Visual Studio.</li>
<li>Configure the manifest file to support task pane integration in PowerPoint.</li>
</ul>
</li>
<li>
<p><strong>UI Conversion</strong></p>
<ul>
<li><strong>Task Pane Interface:</strong><br>
Re-create the form as an HTML page to be loaded in a task pane.</li>
<li><strong>Controls:</strong>
<ul>
<li>Replace  <code>DateTimePicker</code>  with an HTML  <code>&lt;input type="date"&gt;</code>.</li>
<li>Replace radio buttons with standard HTML radio buttons.</li>
<li>Implement buttons (Ok and Close) using HTML  <code>&lt;button&gt;</code>  elements.</li>
</ul>
</li>
</ul>
</li>
<li>
<p><strong>Client-Side Logic</strong></p>
<ul>
<li><strong>Date Format Logic:</strong><br>
Implement JavaScript that listens to radio button change events and updates the  <code>&lt;input type="date"&gt;</code>  display format accordingly. Note: Formatting might require additional JavaScript logic or third-party libraries to achieve non-default formats.</li>
<li><strong>Regional Settings:</strong><br>
Use  <code>Intl.DateTimeFormat</code>  to handle regional formatting preferences.</li>
</ul>
</li>
<li>
<p><strong>PowerPoint Integration</strong></p>
<ul>
<li><strong>Office.js API:</strong><br>
Use the Office.js PowerPoint API (or a compatible approach) to access and update the cover slide.</li>
<li><strong>Cover Update:</strong><br>
Replace the  <code>Cover.setCoverPageDate</code>  call with a JavaScript function that performs similar actions through Office.js commands.</li>
</ul>
</li>
<li>
<p><strong>Event Handling and Communications</strong></p>
<ul>
<li>Use Office.js event bindings to react to user actions (button clicks, radio changes) and update document content.</li>
</ul>
</li>
<li>
<p><strong>Testing</strong></p>
<ul>
<li>Test across platforms supported by Office.js (Windows, Mac, and Office on the web) to ensure consistent behavior of the add-in.</li>
</ul>
</li>
<li>
<p><strong>Documentation</strong></p>
<ul>
<li>Update both internal documentation and user guides with details of the new Office.js-based implementation.</li>
</ul>
</li>
</ol>
<hr>
<h1 id="task-list-and-estimate">Task List and Estimate</h1>
<ul>
<li>
<p><strong>Project Setup (4 hours)</strong></p>
<ul>
<li>Create the basic Office.js add-in project.</li>
<li>Configure manifest and task pane integration.</li>
</ul>
</li>
<li>
<p><strong>UI Development (8 hours)</strong></p>
<ul>
<li>Develop the HTML/CSS interface:
<ul>
<li>Date picker</li>
<li>Radio buttons for format selection</li>
<li>Ok and Close buttons</li>
</ul>
</li>
</ul>
</li>
<li>
<p><strong>Logic Implementation (JavaScript) (6 hours)</strong></p>
<ul>
<li>Implement date format update logic.</li>
<li>Apply regional settings using  <code>Intl.DateTimeFormat</code>.</li>
</ul>
</li>
<li>
<p><strong>PowerPoint Integration (6 hours)</strong></p>
<ul>
<li>Replace VSTO  <code>Cover.setCoverPageDate</code>  with Office.js-based update functionality.</li>
<li>Develop and test integration with PowerPoint slides.</li>
</ul>
</li>
<li>
<p><strong>Testing and Debugging (6 hours)</strong></p>
<ul>
<li>Cross-platform testing on supported clients.</li>
<li>Debug browser and API issues.</li>
</ul>
</li>
<li>
<p><strong>Documentation &amp; Training (4 hours)</strong></p>
<ul>
<li>Update user/developer documentation.</li>
<li>Create migration guidelines and release notes.</li>
</ul>
</li>
</ul>
<p><strong>Total Estimate:</strong>  Approximately 34 hours (4.25 days)</p>
<pre><code>// This is a single file React component that displays the UI only.
import React from "react";

const CoverDateChange = () =&gt; {
  return (
    &lt;div className="cover-date-change"&gt;
      &lt;div className="group-box"&gt;
        
        &lt;input type="date" className="date-picker" /&gt;
        &lt;div className="radio-group"&gt;
          &lt;label&gt;
            &lt;input
              type="radio"
              name="dateFormat"
              value="month-year"
              defaultChecked
            /&gt;{" "}
            Month Year
          &lt;/label&gt;
          &lt;label&gt;
            &lt;input type="radio" name="dateFormat" value="month-day-year" /&gt;{" "}
            Month Day, Year
          &lt;/label&gt;
          &lt;label&gt;
            &lt;input type="radio" name="dateFormat" value="day-month-year-eu" /&gt;{" "}
            Day Month, Year (EU)
          &lt;/label&gt;
          &lt;label&gt;
            &lt;input type="radio" name="dateFormat" value="no-date" /&gt; No Date
          &lt;/label&gt;
        &lt;/div&gt;
      &lt;/div&gt;
      &lt;div className="button-group"&gt;
        &lt;button className="btn ok"&gt;Ok&lt;/button&gt;
        &lt;button className="btn close"&gt;Close&lt;/button&gt;
      &lt;/div&gt;
      &lt;style&gt;{`
        .cover-date-change {
          font-family: Arial, sans-serif;
          width: 428px;
          margin: 20px auto;
          padding: 20px;
          border: 1px solid #ccc;
          border-radius: 4px;
          background-color: #fff;
        }
        .group-box {
          border: 1px solid #ddd;
          padding: 15px;
          margin-bottom: 20px;
          position: relative;
        }
        .picture {
          position: absolute;
          left: 15px;
          top: 15px;
          width: 48px;
          height: 48px;
        }
        .date-picker {
          margin-left: 80px;
          padding: 5px;
          font-size: 14px;
          width: 200px;
          margin-bottom: 20px;
        }
        .radio-group {
          margin-left: 15px;
        }
        .radio-group label {
          display: block;
          margin-bottom: 10px;
          font-size: 14px;
          cursor: pointer;
        }
        .button-group {
          text-align: right;
        }
        .btn {
          background-color: #e1e1e1;
          color: #000;
          border: 1px solid silver;
          padding: 5px 15px;
          margin-left: 5px;
          font-size: 14px;
          cursor: pointer;
        }
      `}&lt;/style&gt;
    &lt;/div&gt;
  );
};

export default CoverDateChange;
</code></pre>

