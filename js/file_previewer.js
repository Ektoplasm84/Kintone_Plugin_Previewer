function checkFileExtension(file) {
  var excelReg = /\.xl(s[xmb]|t[xm]|am|s)$/g; // Excel file extension regex
  var pdfReg = /\.pdf$/g; // PDF file extension regex
  return excelReg.test(file.name) || pdfReg.test(file.name);
}

// Style Overrides (fixes strange behaviour on underlines)

function addAnchorOverride() {
  var styleElement = document.createElement("style");
  styleElement.innerHTML =
    "a { text-decoration: none !important; } " +
    "h2 { all: unset !important; font-size: inherit; color: inherit; }";
  document.head.appendChild(styleElement);
}

function loadExcelModal(fileInfo) {
  console.log("Loading modal for file:", fileInfo.name);
  console.log("File Key:", fileInfo.fileKey);
  addAnchorOverride();

  var previewElement;
  if (window.innerWidth > 768) {
    // Desktop preview
    jQuery(".file-image-container-gaia").each(function (i, e) {
      var fileName = jQuery(e).children("a:eq(0)").text();
      if (
        fileName == fileInfo.name &&
        jQuery(e).children(".gaia-argoui-button").length == 0
      ) {
        previewElement = jQuery(e);
        return false;
      }
    });
  } else {
    // Mobile preview
    jQuery(".control-showlayout-file-gaia .control-value-gaia div").each(
      function (index, e) {
        var fileName = jQuery(e).find("a:eq(0)").text();
        if (
          fileName == fileInfo.name &&
          jQuery(e).find(".gaia-argoui-button").length == 0
        ) {
          previewElement = jQuery(e);
          return false;
        }
      }
    );
  }

  if (!previewElement) {
    console.log(
      "Element with class 'file-image-container-gaia' not found or button already added."
    );
    return;
  }

  var recordGaiaElement = jQuery("#record-gaia");
  var modalId = "MyModal" + fileInfo.fileKey;
  var tabId = "myTab" + fileInfo.fileKey;
  var tabContentId = "tab-content" + fileInfo.fileKey;
  var $button = jQuery(
    '<button type="button" class="gaia-argoui-button" data-toggle="modal" data-target="#' +
      modalId +
      '"><i class="fa fa-search"></i></button>'
  );

  $button.on("click", function () {
    console.log("Button clicked!");
    console.log("Trying to load:", fileInfo.name);
    jQuery("#" + modalId).modal("show");
  });

  var MyModal = jQuery(
    '<div class="modal fade tab-pane active" id="' +
      modalId +
      '" tabindex="-1" role="dialog" aria-labelledby="MyModalLabel">' +
      '<div class="modal-dialog ' +
      (window.innerWidth > 768 ? "modal-xl" : "modal-dialog-scrollable") +
      '" style="border-radius:5px" role="document">' +
      '<div class="modal-content">' +
      '<div class="modal-header">' +
      '<h5 class="modal-title">' +
      fileInfo.name +
      "</h5>" +
      '<button type="button" class="close" data-dismiss="modal" aria-label="Close">' +
      '<span aria-hidden="true">&times;</span>' +
      "</button>" +
      "</div>" +
      '<ul class="nav nav-tabs" id=' +
      tabId +
      ">" +
      "</ul>" +
      "<div id=" +
      tabContentId +
      ' class="tab-content">' +
      '<div class="d-flex justify-content-center">' +
      '<div class="spinner-border" role="status">' +
      '<span class="sr-only">Loading...</span>' +
      "</div>" +
      "</div>" +
      "</div>" +
      '<div class="modal-footer">' +
      '<button type="button" class="btn btn-secondary" data-dismiss="modal">CLOSE</button>' +
      "</div>" +
      "</div>" +
      "</div>" +
      "</div>"
  );

  // Add event listener to the modal close buttons to close the modal
  MyModal.find('[data-dismiss="modal"]').on("click", function () {
    MyModal.modal("hide");
  });

  previewElement.append($button);
  recordGaiaElement.append(MyModal);
  jQuery("body").prepend(MyModal);

  // Add event listener to tab links to load the corresponding sheet content
  jQuery("#" + tabId).on("click", "a.nav-link", function (e) {
    e.preventDefault(); // Prevent the default behavior of the link
    var sheetIndex = jQuery(this).attr("href").split("-").pop();
    loadExcelFile(fileInfo, tabContentId, sheetIndex);
  });

  // Add event listener to the modal shown event
  jQuery("#" + modalId).on("shown.bs.modal", function (e) {
    loadExcelFile(fileInfo, tabContentId, 0); // Load the first sheet when the modal is shown
  });
}

function loadPdfModal(fileInfo) {
  console.log("Loading PDF modal for file:", fileInfo.name);
  var modalId = "PdfModal" + fileInfo.fileKey;
  var pdfUrl = "/k/v1/file.json?fileKey=" + fileInfo.fileKey;

  var $button = jQuery(
    '<button type="button" class="gaia-argoui-button" data-toggle="modal" data-target="#' +
      modalId +
      '"><i class="fa fa-search"></i></button>'
  );

  $button.on("click", function () {
    console.log("Button clicked!");
    console.log("Trying to load:", fileInfo.name);
    loadPdfFile(pdfUrl, modalId);
    jQuery("#" + modalId).modal("show");
  });

  var previewElement;
  if (window.innerWidth > 768) {
    // Desktop preview
    jQuery(".file-image-container-gaia").each(function (i, e) {
      var fileName = jQuery(e).children("a:eq(0)").text();
      if (
        fileName == fileInfo.name &&
        jQuery(e).children(".gaia-argoui-button").length == 0
      ) {
        previewElement = jQuery(e);
        return false;
      }
    });
  } else {
    // Mobile preview
    jQuery(".control-showlayout-file-gaia .control-value-gaia div").each(
      function (index, e) {
        var fileName = jQuery(e).find("a:eq(0)").text();
        if (
          fileName == fileInfo.name &&
          jQuery(e).find(".gaia-argoui-button").length == 0
        ) {
          previewElement = jQuery(e);
          return false;
        }
      }
    );
  }

  if (!previewElement) {
    console.log(
      "Element with class 'file-image-container-gaia' not found or button already added."
    );
    return;
  }

  var recordGaiaElement = jQuery("#record-gaia");
  var MyModal = jQuery(
    '<style type="text/css">iframe { border: none; width: 100%; height: 80vh; }</style>' +
      '<div class="modal fade" id="' +
      modalId +
      '" tabindex="-1" role="dialog" aria-labelledby="PdfModalLabel">' +
      '<div class="modal-dialog modal-xl" style="border-radius: 5px;" role="document">' +
      '<div class="modal-content">' +
      '<div class="modal-header">' +
      '<h5 class="modal-title">' +
      fileInfo.name +
      "</h5>" +
      '<button type="button" class="close" data-dismiss="modal" aria-label="Close">' +
      '<span aria-hidden="true">&times;</span>' +
      "</button>" +
      "</div>" +
      '<div class="modal-body">' +
      '<div id="' +
      modalId +
      'PdfContainer"></div>' + // Use unique ID for each pdfContainer
      "</div>" +
      '<div class="modal-footer">' +
      '<button type="button" class="btn btn-secondary" data-dismiss="modal">CLOSE</button>' +
      "</div>" +
      "</div>" +
      "</div>" +
      "</div>"
  );

  // Add event listener to the modal close buttons to close the modal
  MyModal.find('[data-dismiss="modal"]').on("click", function () {
    console.log("Modal closed!");
    MyModal.modal("hide");
  });

  console.log("Adding button to the preview element.");
  previewElement.append($button);
  recordGaiaElement.append(MyModal);
  jQuery("body").prepend(MyModal);
  addAnchorOverride();
}

function loadExcelFile(fileInfo, tabContentId, sheetIndex) {
  var fileUrl = "/k/v1/file.json?fileKey=" + fileInfo.fileKey;
  readWorkbookFromRemoteFile(fileUrl, function (workbook) {
    readWorkbook(workbook, fileInfo, tabContentId, sheetIndex);
  });
}

function readWorkbookFromRemoteFile(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open("get", url, true);
  xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");
  xhr.responseType = "arraybuffer";
  xhr.onload = function (e) {
    if (xhr.status == 200) {
      var data = new Uint8Array(xhr.response);
      var workbook = XLSX.read(data, { type: "array" });
      if (callback) callback(workbook);
    }
  };
  xhr.send();
}

function readWorkbook(workbook, fileInfo, tabContentId, sheetIndex) {
  var sheetNames = workbook.SheetNames;
  var navHtml = "";
  var tabHtml = "";
  var myTabId = "myTab" + fileInfo.fileKey;

  for (var index = 0; index < sheetNames.length; index++) {
    var sheetName = sheetNames[index];
    var worksheet = workbook.Sheets[sheetName];
    var sheetHtml = XLSX.utils.sheet_to_html(worksheet);

    // Remove the underlining style from the sheetHtml
    sheetHtml = sheetHtml.replace(/<style.*<\/style>/, "");

    var tabid = "tab" + fileInfo.fileKey + "-" + index;
    var xlsid = "xlsid" + fileInfo.fileKey + "-" + index;
    var active = index == sheetIndex ? "active" : "";

    navHtml +=
      '<li class="nav-item"><a class="nav-link ' +
      active +
      '" href="#' +
      tabid +
      '" data-toggle="tab">' +
      sheetName +
      "</a></li>";

    tabHtml +=
      '<div id="' +
      tabid +
      '" class="tab-pane ' +
      active +
      '" style="padding:10px;overflow:auto;height:600px" >' +
      '<div id="' +
      xlsid +
      '">' +
      sheetHtml +
      " </div></div>";
  }

  jQuery("#" + myTabId).html(navHtml);
  jQuery("#" + tabContentId).html(tabHtml);
  jQuery("#" + tabContentId + " table").addClass(
    "table table-bordered table-hover"
  );
}

function loadPdfFile(pdfUrl, modalId) {
  console.log("Loading PDF file from URL:", pdfUrl);

  var xhr = new XMLHttpRequest();
  xhr.open("get", pdfUrl, true);
  xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest"); // Add the required header
  xhr.responseType = "arraybuffer";
  xhr.onload = function (e) {
    if (xhr.status === 200) {
      var data = new Uint8Array(xhr.response);
      var pdfBlob = new Blob([data], { type: "application/pdf" });
      var pdfUrl = URL.createObjectURL(pdfBlob);

      var pdfContainer = document.getElementById(modalId + "PdfContainer"); // Get the corresponding pdfContainer
      pdfContainer.innerHTML = '<iframe src="' + pdfUrl + '"></iframe>';
    }
  };
  xhr.send();
}

kintone.events.on("app.record.index.show", function (event) {
  addAnchorOverride();
});

kintone.events.on("app.record.detail.show", function (event) {
  var record = event.record;
  for (var index in record) {
    var field = record[index];
    if (field.type === "FILE") {
      var fieldValue = field.value;
      fieldValue.forEach(function (file) {
        if (checkFileExtension(file)) {
          if (file.name.endsWith(".pdf")) {
            loadPdfModal(file);
          } else {
            loadExcelModal(file);
          }
        }
        addAnchorOverride();
      });
    }
  }
});
