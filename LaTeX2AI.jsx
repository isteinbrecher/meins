
//////////////////////////////////////////////////////////////////////////////
// UI                                                                       //
//////////////////////////////////////////////////////////////////////////////
try {
    

    var l2a_path = 'C:\\Users\\VMware\\Desktop\\illustrator test\\LaTeX2AI';
    var template_path = 'C:\\LaTeX2AI\\template\\template.ai';
    var temp_path = Folder.temp.fsName;
    l2a_document = new L2A_Document();
    
    // check files and folderstructure
    //check_structure()
    
    // Create UI
    var dlg = new Window('dialog', 'LaTeX2AI');
    dlg.frameLocation = [100,100];
    dlg.size = [250,160];

    // define text
    //dlg.intro = dlg.add('statictext', [20,20,150,40] );
    //dlg.intro.text = 'LaTeX2AI';

    // define buttons
    dlg.create = dlg.add('button', [0,0,250,40], 'new LaTeX2AI item');
    dlg.change = dlg.add('button', [0,40,250,80], 'change selected LaTeX2AI item');
    dlg.box = dlg.add('button', [0,80,250,120], 'redo all LaTeX2AI boundary boxes');
    dlg.redo = dlg.add('button', [0,120,250,160], 'redo all LaTeX2AI elements');

    // add events
    dlg.create.addEventListener ('click', create_l2a );
    dlg.change.addEventListener ('click', change_l2a );
    dlg.redo.addEventListener ('click', redo_l2a );
    dlg.box.addEventListener ('click', redo_bound_l2a );
    dlg.show();
    
}
catch(e) {
    alert( e.message, 'LaTeX2AI Error', true);
}


//////////////////////////////////////////////////////////////////////////////
// Main functions                                                           //
//////////////////////////////////////////////////////////////////////////////

// creates new LaTeX2AI element
function create_l2a( ) {
    try {
        l2a_document.l2a_create();
        app.redraw(); // redraw the illustrator window
    }
    catch(e) {
        alert( e.message, 'LaTeX2AI Error', true);
    }    
}

// creates new LaTeX2AI element
function change_l2a( ) {
    try {
        l2a_document.l2a_change();
        app.redraw(); // redraw the illustrator window
    }
    catch(e) {
        alert( e.message, 'LaTeX2AI Error', true);
    }    
}

// removes all pdfs and redos every element
function redo_l2a( ) {    
    try {
        l2a_document.l2a_redo();
        app.redraw(); // redraw the illustrator window
    }
    catch(e) {
        alert( e.message, 'LaTeX2AI Error', true);
    }    
}

// redos selected boundary boxes
function redo_bound_l2a( ) {    
    try {
        l2a_document.l2a_redo_bound();
        app.redraw(); // redraw the illustrator window
    }
    catch(e) {
        alert( e.message, 'LaTeX2AI Error', true);
    }    
}


//////////////////////////////////////////////////////////////////////////////
// Classes                                                                  //
//////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////
// Class for the document to be used with LaTeX2AI.
///////////////////////////////////////////////////////////
function L2A_Document(){

    
    
    // Creates a new LaTeX2AI object.
    this.l2a_create = function(){
        new L2A_Element();
    }


    // Changes selectet LaTeX2AI objects.
    this.l2a_change = function(){
        
        this.load_l2a_elements();
            
        if (this.l2a_elements_selected.length == 0) {
            throw new Error('No LaTeX2AI elements selected!');
        }
    
        for ( i = 0; i < this.l2a_elements_selected.length; i++ ) {
            this.l2a_elements_selected[i].l2a_change_element();
        }
    }


    // Redos all LaTeX2AI elements.
    this.l2a_redo = function(){

        this.load_l2a_elements();
            
        if (this.l2a_elements.length == 0) {
            throw new Error('No LaTeX2AI elements!');
        }
    
        this.delete_pdf();
    
        for ( i = 0; i < this.l2a_elements.length; i++ ) {
            this.l2a_elements[i].l2a_redo_element();
        }
    }


    // Changes selectet LaTeX2AI objects.
    this.l2a_redo_bound = function(){
        this.load_l2a_elements();
            
        if (this.l2a_elements.length == 0) {
            throw new Error('No LaTeX2AI elements!');
        }
    
        for ( i = 0; i < this.l2a_elements.length; i++ ) {
            this.l2a_elements[i].set_bound();
        }
    }
    
    
    // Fills the object with all LaTeX2AI items in the document.
    this.load_l2a_elements = function(){
        this.l2a_elements = [];
        this.l2a_elements_selected = [];
        
        // loop over all placed items
        for ( i = 0; i < this.document.placedItems.length; i++ ) {
            if (this.document.placedItems[i].name.indexOf('LaTeX2AI_') > -1) {
                this.l2a_elements.push(new L2A_Element(this.document.placedItems[i]));
                if (this.document.placedItems[i].selected == true){
                    this.l2a_elements_selected.push(new L2A_Element(this.document.placedItems[i]));
                }
            }
        }
    }


    // Deletes all pdf files for the current document.
    this.delete_pdf = function() {

        // local variables
        var i, l2a_files

        // all files corresponding to the current document
        l2a_files = Folder(l2a_path).getFiles(this.file_name + '*_LaTeX2AI_*.pdf');

        // removes all elements
        for ( i = 0; i < l2a_files.length; i++ ) {
            l2a_files[i].remove();
        }
    }


    // Reads an array of strings and set the settings acordingly.
    this.read_settings = function(path) {
        
        var content = read_file(path);
        
        for (var i = 0; i < content.length; i++){
            set = read_setting_line(content[i]);
            
            if (!set[1]){
                throw new Error('Empty value on setting "' + set[0] + '"!');
            }
             
            switch (set[0].toLowerCase()) {
                case 'exe':
                    this.latex_exe = set[1];
                    break;
                case 'arg':
                    this.latex_arg = set[1];
                    break;
                default:
                    throw new Error('Setting "' + set[0] + '" unknown!');
            }
        }
    }


    // Reads an array of strings and set the settings acordingly.
    this.read_last_input = function() {
        
        var content = read_file(temp_path + '\\LaTeX2AI_last_input.txt');
        
        if (content.length > 1){
            this.last_input = content[1];
            this.last_pos = content[0];
        }
    }


    // creates the header file
    this.create_header = function() {
        this.header_file.open('w');
        this.header_file.writeln('\\documentclass[11pt]{scrartcl}');
        this.header_file.writeln('');
        this.header_file.writeln('\\usepackage{amsmath}');
        this.header_file.close();
    }


    ///////////////////////////////////////////////////////////////
    // Initiate Object.
    ///////////////////////////////////////////////////////////////
    
    // check if there is an active document
    if (app.documents.length > 0 ) {
        
        // check if the file has a valid filepath
        if ((' ' + app.documents[0].path + ' ').length < 3) {
            throw new Error('File needs to be saved!');
        }
    
        l2a_path = app.documents[0].path + '\\LaTeX2AI';
        
        // check if template file exists
        this.template_file = File( template_path );
        if (this.template_file.exists == false) {
            throw new Error('Template file does not exist!');
        }
        
        // check if 'LaTeX2AI' folder exists in the active document directory
        if (Folder(l2a_path).exists == false) {
            Folder(l2a_path).create();
            alert('<LaTeX2AI> folder did not exist in document directory, was created!');
        }
        
        // check if header file exists in 'LaTeX2AI'
        this.header_file = File(l2a_path + '\\_LaTeX2AI_header.tex');
        if (this.header_file.exists == false) {
            this.create_header();
            alert('<LaTeX2AI> folder did not contain <_LaTeX2AI_header.tex> file, was created (with font site 11pt)!');
        }
        
        
        this.document = app.documents[0];
        this.file_name = get_filename(this.document.fullName);
        this.latex_exe = 'pdflatex';
        this.latex_arg = '-interaction nonstopmode -halt-on-error -file-line-error';
        this.last_input = '$$';
        this.last_pos = 'cc';
        
        this.read_settings(l2a_path + '\\_LaTeX2AI_settings.txt');
        this.read_last_input();
    
    }
	else{
		throw new Error('No open documents!');
	}    
}



///////////////////////////////////////////////////////////
// Class for a LaTeX2AI element.
///////////////////////////////////////////////////////////
function L2A_Element(placed_item){
    
    
    // Creates the LaTeX pdf and moves it into the \LaTeX2AI folder.
    this.create_latex = function(){

        // remove the pdf file if it exists
        var pdf_file = File(temp_path + '\\' + this.latex_name + '.pdf');
        if (pdf_file.exists == true) {
            pdf_file.remove();
        }
        
        // create and execute the batch file
        var batch_file = this.write_latex();
        batch_file.execute();
        
        // wait until batch file is deleted
        while (batch_file.exists == true) {
        }
    
        // check if the LaTeX file was compilled propperly
        if (pdf_file.exists == false) {
            throw new Error('No .pdf file produced! LaTeX error?');
        }
        else{
            // create the pdf file object and move it to the latex2AI folder
            pdf_file.copy (l2a_path + '\\' + this.latex_name + '.pdf');
            this.pdf_file = File(l2a_path + '\\' + this.latex_name + '.pdf');
        }
        
    }


    // Writes the LaTeX files into the temp directory.
    this.write_latex = function(){

        // copies the header file into the temp directory
        l2a_document.header_file.copy(temp_path + '\\_LaTeX2AI_header.tex')
        
        // creates the latex file in the temp folder
        var latex_file = File(temp_path + '\\' + this.latex_name + '.tex');
        latex_file.open('w');
        latex_file.writeln('\\input{_LaTeX2AI_header}');
        latex_file.writeln('\\usepackage[active,tightpage]{preview}');
        latex_file.writeln('\\usepackage{varwidth}');
        latex_file.writeln('\\AtBeginDocument{\\begin{preview}\\begin{varwidth}{\\linewidth}}');
        latex_file.writeln('\\AtEndDocument{\\end{varwidth}\\end{preview}}');
        latex_file.writeln('\\pagestyle{empty}');
        latex_file.writeln('\\begin{document}');
        latex_file.writeln(this.latex_input);
        latex_file.writeln('\\end{document}');
        latex_file.close();
        
        // creates the batch file
        var batch_file = File(temp_path + '\\' + 'LaTeX2AI_batch.bat');
        batch_file.open('w');
        batch_file.writeln('cd "' + temp_path + '"' );
        batch_file.writeln('"' + l2a_document.latex_exe + '" ' + l2a_document.latex_arg + ' "' + this.latex_name + '.tex"' );
        batch_file.writeln('del "' + temp_path +'\\LaTeX2AI_batch.bat"');
        batch_file.close();
        
        return batch_file;
    }


    // Writes the last input to the Temp folder.
    this.write_last_input = function(){
        var last_input = File(temp_path + '\\LaTeX2AI_last_input.txt');
        last_input.open('w');
        last_input.writeln(this.pos_input);
        last_input.writeln(this.latex_input);
        last_input.close();
    }


    // Gets the filename for the next LaTeX2AI item.
    this.set_l2a_name = function () {
        // loop to find first nonused name for LaTeX2AI element
        var i = 1;
        while (File(l2a_path + '\\' + l2a_document.file_name + '_LaTeX2AI_' + pad(i,3) + '.pdf').exists) {
            i = i+1;
        }
        this.latex_name = l2a_document.file_name + '_LaTeX2AI_' + pad(i,3);
    }


    // Get the LaTeX code for the LaTeX2AI element.
    this.get_latex_code_input = function() {

        if (typeof this.latex_input == 'undefined'){
            this.latex_input = l2a_document.last_input;
        }

        // input form for LaTeX code
        this.latex_input = prompt('Please enter the LaTeX code', this.latex_input, 'LaTeX2AI');
        
        // check if there is an input
        if(this.latex_input==null) {
            throw new Error('No LaTeX code!');
        }
    
        // update the last input for the document
        l2a_document.last_input = this.latex_input;
    }

    // Get the position input for the LaTeX2AI element.
    this.get_pos_input = function() {
        // input form for LaTeX positioning
        this.pos_input = prompt('Please enter positioning of the text', l2a_document.last_pos,'LaTeX2AI');
        
        // check if there position input is valid
        this.get_posfac();
        
        // update the last input for the document
        l2a_document.last_pos = this.pos_input;
    }

    
    // Gets the position factor array from pos input string.
    this.get_posfac = function() {
        
        var pos;

        switch (this.pos_input) {
            case 'tl':
                pos = [0.,0.];
                break;
            case 'tc':
                pos = [0.5,0];
                break;
            case 'tr':
                pos = [1.,0.];
                break;
            case 'cl':
                pos = [0.,0.5];
                break;
            case 'cc':
                pos = [0.5,0.5];
                break;
            case 'cr':
                pos = [1.,0.5];
                break;
            case 'bl':
                pos = [0.,1.];
                break;
            case 'bc':
                pos = [0.5,1.];
                break;
            case 'br':
                pos = [1.,1.];
                break;
            default:
                throw new Error('Wrong input for the positioning!');
        }
        this.posfac = pos;
    }

    
    // Gets relevant placed item from template document.
    this.get_template = function(){
        
        // open the template file
        var template_doc = app.open(l2a_document.template_file);
        
        // copy the placedItem to the destination doc and close the template doc
        var template_item = template_doc.placedItems.getByName ('LaTeX2AI_temp_' + this.pos_input);
        template_item.selected = true;
        app.copy();
        template_doc.close(SaveOptions.DONOTSAVECHANGES);
        l2a_document.document.activate();
        app.paste();
                
        // finish the item
        this.placed_item = l2a_document.document.placedItems.getByName('LaTeX2AI_temp_' + this.pos_input);
        this.relink_item();
        this.placed_item.selected = true;
        
    }

    // Relinks the pdf file and updates the the note attribute of the placed item.
    this.relink_item = function(){
        this.placed_item.file = this.pdf_file;
        this.set_attributes();
        
    }


    // Updates the the note attribute of the placed item.    
    this.set_attributes = function(){
        this.placed_item.name = this.latex_name;
        this.placed_item.note = this.pos_input + this.latex_input;
        
    }


    // Sets the boundary box for the LaTeX2AI element.
    this.set_bound = function() {
        var err = 0.001; // max error for equality to be detected
        
        // only do something if object is not rotated
        if ( Math.abs(this.placed_item.matrix.mValueB) < err){
        
            // variables
            var M, X; // matrix
            var height, width, height_new, width_new; // numbers
            
            // get the original pdf size
            M = this.placed_item.boundingBox;
            width_new = Math.abs(M[0]-M[2]);
            height_new = Math.abs(M[1]-M[3]);
            
            // get the actual position height and width
            width = this.placed_item.width;
            height = this.placed_item.height;
            X = this.placed_item.position;
            
            // check if AI box is smaller than pdf
            if (Math.abs(width - width_new) < err || Math.abs(height - height_new) < err ) {
                this.placed_item.width = 2*width_new;
                this.placed_item.height = 2*height_new;
            }
            
            // set to new values
            this.placed_item.width = width_new;    
            this.placed_item.height = height_new;
            this.placed_item.position = [X[0] + this.posfac[0] * (width - width_new) , X[1] - this.posfac[1] * (height - height_new)];
        }
    }


    // Changes the l2a element
    this.l2a_change_element = function() {
        var old_input = this.latex_input;
        this.get_latex_code_input();
        if (old_input !== this.latex_input){
            this.create_latex();
            this.relink_item();
            this.set_bound();
        }
    }


    // Redos the l2a element
    this.l2a_redo_element = function() {
        this.set_l2a_name();
        this.create_latex();
        this.relink_item();
        this.set_bound();
    }


    ///////////////////////////////////////////////////////////////
    // Initiate Object.
    ///////////////////////////////////////////////////////////////
    
    // if object is initialised with placed item the item is added
    if (typeof placed_item !== 'undefined'){
        this.placed_item = placed_item;
        this.latex_name = placed_item.name;
        this.pdf_file;
        this.pos_input = placed_item.note.substr(0,2);
        this.get_posfac();
        this.latex_input = placed_item.note.substr(2,placed_item.note.length);
    }
    // otherwise the input is retrieved and the object is created
    else{
        this.set_l2a_name();        
        this.get_pos_input();        
        this.get_latex_code_input();
        this.create_latex();
        this.get_template();
        this.set_bound();
        this.write_last_input();
    }
    
}


//////////////////////////////////////////////////////////////////////////////
// functions                                                                //
//////////////////////////////////////////////////////////////////////////////


// Returns the filename of an File object (without the extension).
function get_filename( full_name ) {
    
    var path = 'x' + full_name;
    var a = path.lastIndexOf('/');
    var b = path.lastIndexOf('.');
    return path.substr(a+1,b-a-1);
}


// Padds the number num with zeros so a total length of size is reached.
function pad(num, size) {
    var s = num+"";
    while (s.length < size) s = "0" + s;
    return s;
}


// Reads a file and returns a dictionary with the settings and values.
function read_file(path) {

    var comment = "%";
    var file = File(path);
    var content = [];
    
    if (file.exists){
                
        file.open("r");
        
        while(! file.eof){
            var line = file.readln();
            if (line.length == 0){
                continue;
            }
            var index = line.indexOf(comment);
            if (index == -1){
                content.push(line);
            }
            else if (index > 0){
                content.push(line.substring(0, index));
            }
        }
    }
    file.close();
    return content;
}


// Reads a settings line and returns the key and value.
function read_setting_line(line){
    var index = line.indexOf("=");
    var key = line.substring(0, index);
    var value = line.substring(index+1, line.length);    
    return [key.trim(), value.trim()];
}


// Returns the keys of a dictionary.
function get_keys(dict){
    var keys = [];
    for (var key in dict){
         keys.push(key);
    }
    return keys;
}


// Add trim method to strings.
if (!String.prototype.trim) {
    String.prototype.trim = function() {
        return this.replace(/^\s+|\s+$/g,'');
    }
}