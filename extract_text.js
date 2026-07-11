const mammoth = require("mammoth");

mammoth.extractRawText({path: "Vimlesh_GitHub_Projects_Summary.docx"})
    .then(function(result){
        var text = result.value; 
        console.log(text);
    })
    .done();
