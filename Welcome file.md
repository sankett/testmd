---


---

<h3 id="explanation-of-existing-form-functionality">Explanation of Existing Form Functionality</h3>
<p>The provided <a href="http://VB.NET">VB.NET</a> form  <code>frm_Chapter</code>  is part of a VSTO (Visual Studio Tools for Office) add-in for PowerPoint. It is designed to allow users to insert chapters or subchapters into a PowerPoint presentation. The form provides the following functionalities:</p>
<ol>
<li><strong>Chapter/Subchapter Selection</strong>: Users can choose between inserting a chapter or a subchapter using a dropdown (<code>cmbChapter</code>).</li>
<li><strong>Preview and Thumbnail Gallery</strong>: Displays a gallery of banner images for chapters/subchapters. Users can select a banner image for the chapter.</li>
<li><strong>Pagination</strong>: Handles pagination for the gallery, allowing users to navigate through multiple pages of banner images.</li>
<li><strong>Dynamic Data Loading</strong>: Fetches banner images dynamically from a database or file system.</li>
<li><strong>PowerPoint Integration</strong>: Inserts a new slide into the active PowerPoint presentation with the selected chapter/subchapter details and banner image.</li>
<li><strong>Validation</strong>: Ensures that subchapters can only be added if a preceding chapter exists.</li>
<li><strong>Progress Indicator</strong>: Displays a progress bar while loading banner images.</li>
<li><strong>Help Button</strong>: Provides a link to an online help guide.</li>
</ol>
<hr>
<h3 id="list-of-methods-events-and-properties">List of Methods, Events, and Properties</h3>
<h4 id="methods"><strong>Methods</strong></h4>
<ol>
<li><strong>SetBannerCategoryDictionaryFromCoverPage</strong>: Initializes the banner images dictionary based on the cover page category.</li>
<li><strong>frmChapter_Load</strong>: Handles form load events, initializes controls, and sets up data.</li>
<li><strong>cmbChapter_SelectedIndexChanged</strong>: Handles changes in the chapter/subchapter dropdown selection.</li>
<li><strong>btnOk_Click</strong>: Handles the “Insert” button click event to insert a chapter/subchapter into the PowerPoint presentation.</li>
<li><strong>HasPrecedingChapter</strong>: Checks if a preceding chapter exists in the presentation.</li>
<li><strong>GetPrecedingChapterIndex</strong>: Retrieves the index of the preceding chapter.</li>
<li><strong>GetPrecedingChapterTagBanner</strong>: Retrieves the banner tag of the preceding chapter.</li>
<li><strong>PreparePreviewAndThumbnailGallery</strong>: Loads thumbnails for the banner gallery.</li>
<li><strong>downloadingDataGridView_CellPainting</strong>: Customizes the appearance of the DataGridView cells.</li>
<li><strong>BackgroundWorker1_DoWork</strong>: Handles background loading of banner images.</li>
<li><strong>BackgroundWorker1_ProgressChanged</strong>: Updates the progress bar during background work.</li>
<li><strong>BackgroundWorker1_RunWorkerCompleted</strong>: Finalizes the progress bar after background work is complete.</li>
<li><strong>TraverseBannerImagesDictionaryForThumbnailPreviews</strong>: Populates the DataGridView with banner thumbnails.</li>
<li><strong>Grid_CellClick</strong>: Handles cell click events in the DataGridView.</li>
<li><strong>AvailableOptions</strong>: Configures available options based on the current state of the presentation.</li>
<li><strong>FormHelpButtonClicked</strong>: Opens the help guide in a browser.</li>
<li><strong>CreatePagination</strong>: Creates pagination controls for the gallery.</li>
<li><strong>PaginationLabel_Click</strong>: Handles pagination label click events.</li>
<li><strong>PaginationLabel_MouseEnter</strong>: Handles mouse hover events for pagination labels.</li>
<li><strong>PaginationLabel_MouseLeave</strong>: Handles mouse leave events for pagination labels.</li>
</ol>
<h4 id="events"><strong>Events</strong></h4>
<ol>
<li><strong>Load</strong>: Triggered when the form is loaded.</li>
<li><strong>cmbChapter.SelectedIndexChanged</strong>: Triggered when the chapter/subchapter dropdown selection changes.</li>
<li><strong>btnOk.Click</strong>: Triggered when the “Insert” button is clicked.</li>
<li><strong>BackgroundWorker1.DoWork</strong>: Triggered when the background worker starts.</li>
<li><strong>BackgroundWorker1.ProgressChanged</strong>: Triggered when the background worker reports progress.</li>
<li><strong>BackgroundWorker1.RunWorkerCompleted</strong>: Triggered when the background worker completes.</li>
<li><strong>DataGridView.CellClick</strong>: Triggered when a cell in the DataGridView is clicked.</li>
<li><strong>HelpButtonClicked</strong>: Triggered when the help button is clicked.</li>
</ol>
<h4 id="properties"><strong>Properties</strong></h4>
<ol>
<li><strong>App</strong>: Represents the PowerPoint application.</li>
<li><strong>presActive</strong>: Represents the active PowerPoint presentation.</li>
<li><strong>dru</strong>: Instance of  <code>DpiRuntimeUtil</code>  for DPI scaling.</li>
<li><strong>dpiFactor</strong>: DPI scaling factor.</li>
<li><strong>CategoryFilterTriggeredByUser</strong>: Tracks if the category filter was triggered by the user.</li>
<li><strong>PaginationTriggeredByUser</strong>: Tracks if pagination was triggered by the user.</li>
<li><strong>PrevLabelClicked</strong>: Tracks the previously clicked pagination label.</li>
<li><strong>totalRec</strong>: Total number of records available.</li>
<li><strong>pageSize</strong>: Number of records displayed per page.</li>
<li><strong>noOfPages</strong>: Total number of pages available.</li>
<li><strong>currPage</strong>: Current page number.</li>
<li><strong>startPage</strong>: Starting page number for pagination.</li>
<li><strong>originalForeColorOfFont</strong>: Original font color for pagination labels.</li>
<li><strong>paginationBtnHoverColor</strong>: Hover color for pagination labels.</li>
<li><strong>selectedCoverPage</strong>: Index of the selected cover page.</li>
<li><strong>placeholderBitmap</strong>: Placeholder image for empty cells.</li>
<li><strong>previewThumbnailFilePaths</strong>: File paths for preview thumbnails.</li>
<li><strong>coverPageJPGFilePaths</strong>: File paths for cover page images.</li>
<li><strong>PrevSelectedCell</strong>: Tracks the previously selected cell in the DataGridView.</li>
<li><strong>BannerImagesDictionary</strong>: Dictionary of banner images.</li>
<li><strong>BannerImageCategories</strong>: Dictionary of banner image categories.</li>
<li><strong>recordsOffset</strong>: Offset for database records.</li>
<li><strong>theCoverPageExists</strong>: Indicates if the cover page exists.</li>
<li><strong>precedingChapterIndex</strong>: Index of the preceding chapter.</li>
<li><strong>precedingChapterTagBanner</strong>: Tag banner of the preceding chapter.</li>
<li><strong>currentSlideIndex</strong>: Index of the current slide.</li>
<li><strong>selectedChapterImage</strong>: ID of the selected chapter image.</li>
<li><strong>myPitchbookFlypageBannerImages</strong>: Instance of  <code>PitchbookFlypageBannerImages</code>.</li>
</ol>
<hr>
<h3 id="connectionsdependencies-with-modules-and-classes">Connections/Dependencies with Modules and Classes</h3>
<ol>
<li><strong>PowerPoint Integration</strong>: Uses  <code>Microsoft.Office.Interop.PowerPoint</code>  to interact with PowerPoint presentations.</li>
<li><strong>Database Access</strong>: Relies on  <code>PitchbookFlypageBannerImages</code>  to fetch banner images and categories from a database.</li>
<li><strong>UI Components</strong>: Extensively uses Windows Forms controls like  <code>DataGridView</code>,  <code>ComboBox</code>,  <code>PictureBox</code>, and  <code>Label</code>.</li>
<li><strong>Utility Classes</strong>:
<ul>
<li><code>DpiRuntimeUtil</code>: For DPI scaling.</li>
<li><code>UtilityProcedures</code>: For utility functions like  <code>GetCoverPageFormat</code>.</li>
</ul>
</li>
<li><strong>BackgroundWorker</strong>: For asynchronous loading of banner images.</li>
<li><strong>Constants</strong>: Uses constants like  <code>TAG_BANNER</code>,  <code>TAG_FLYPAGE_TYPE</code>, and  <code>TAG_FLYPAGE_NAME</code>  for tagging slides.</li>
</ol>
<hr>
<h3 id="implementation-plan-for-migrating-to-office.js">Implementation Plan for Migrating to Office.js</h3>
<h4 id="step-1-analyze-requirements"><strong>Step 1: Analyze Requirements</strong></h4>
<ul>
<li>Identify all functionalities of the existing form.</li>
<li>Map equivalent Office.js APIs for PowerPoint integration.</li>
</ul>
<h4 id="step-2-set-up-office.js-project"><strong>Step 2: Set Up Office.js Project</strong></h4>
<ul>
<li>Create a new Office Add-in project using React.</li>
<li>Configure the manifest file for PowerPoint.</li>
</ul>
<h4 id="step-3-design-react-component"><strong>Step 3: Design React Component</strong></h4>
<ul>
<li>Create a React component for the form.</li>
<li>Use Fluent UI for styling and controls.</li>
</ul>
<h4 id="step-4-implement-functionalities"><strong>Step 4: Implement Functionalities</strong></h4>
<ol>
<li><strong>Dropdown for Chapter/Subchapter</strong>: Use Fluent UI’s  <code>Dropdown</code>  component.</li>
<li><strong>Thumbnail Gallery</strong>: Use a grid layout with image previews.</li>
<li><strong>Pagination</strong>: Implement pagination using React state and Fluent UI’s  <code>Pagination</code>  component.</li>
<li><strong>PowerPoint Integration</strong>: Use Office.js APIs to insert slides and manipulate tags.</li>
<li><strong>Progress Indicator</strong>: Use Fluent UI’s  <code>ProgressIndicator</code>  component.</li>
<li><strong>Help Button</strong>: Add a button to open the help guide in a browser.</li>
</ol>
<h4 id="step-5-test-and-debug"><strong>Step 5: Test and Debug</strong></h4>
<ul>
<li>Test the add-in in PowerPoint.</li>
<li>Debug any issues and ensure feature parity with the existing form.</li>
</ul>
<h4 id="step-6-deployment"><strong>Step 6: Deployment</strong></h4>
<ul>
<li>Package the add-in and deploy it to the Office Add-ins store or distribute it internally.</li>
</ul>
<hr>
<h3 id="conversion-to-office.js-react-component">Conversion to Office.js React Component</h3>
<h4 id="component-structure"><strong>Component Structure</strong></h4>
<ol>
<li><strong>Form Component</strong>: Main component for the form.</li>
<li><strong>Gallery Component</strong>: Subcomponent for the thumbnail gallery.</li>
<li><strong>Pagination Component</strong>: Subcomponent for pagination controls.</li>
</ol>
<h4 id="key-office.js-apis"><strong>Key Office.js APIs</strong></h4>
<ol>
<li><code>Office.context.document</code>: To interact with the PowerPoint document.</li>
<li><code>Office.context.document.getSelectedDataAsync</code>: To get the current slide.</li>
<li><code>Office.context.document.setSelectedDataAsync</code>: To insert a new slide.</li>
<li><code>Office.context.document.settings</code>: To manage slide tags.</li>
</ol>
<h4 id="implementation-details"><strong>Implementation Details</strong></h4>
<ul>
<li>Use React state to manage form data and gallery state.</li>
<li>Fetch banner images from a REST API or database.</li>
<li>Use Fluent UI for consistent styling.</li>
</ul>
<hr>
<h3 id="estimate-based-on-task">Estimate Based on Task</h3>
<p>|  <strong>Task</strong>  |  <strong>Estimated Hours</strong>  | |-----------------------------------|---------------------| | Analyze existing form | 4 hours | | Set up Office.js project | 6 hours | | Design React component | 8 hours | | Implement dropdown functionality | 4 hours | | Implement gallery functionality | 10 hours | | Implement pagination | 6 hours | | Integrate with PowerPoint (Office.js) | 12 hours | | Add progress indicator | 2 hours | | Add help button | 1 hour | | Test and debug | 8 hours | | Deployment | 4 hours | |  <strong>Total</strong>  |  <strong>65 hours</strong>  |</p>
<hr>
<h3 id="next-steps">Next Steps</h3>
<ol>
<li>Confirm the migration plan and estimate.</li>
<li>Set up the Office.js project.</li>
<li>Begin implementation of the React component.</li>
</ol>

