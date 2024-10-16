<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128611498/23.1.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E3280)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Rich Text Editor for WinForms - Handle the CalculateDocumentVariable Event to Insert Formatted Content

This example illustrates how to use the [Document.CalculateDocumentVariable](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.Document.CalculateDocumentVariable) event to insert formatted content. The [DOCVARIABLE](https://docs.devexpress.com/WindowsForms/9721/controls-and-libraries/rich-text-editor/fields/field-codes/docvariable) has a parameter specified by the [MERGEFIELD](https://docs.devexpress.com/WindowsForms/9714/controls-and-libraries/rich-text-editor/fields/field-codes/mergefield) field and looks as follows:

`{ DOCVARIABLE DetailInfo "{ MERGEFIELD DetailId }" }`

The formatted content is loaded into a non-visual [RichEditDocumentServer](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer) component, which in turn is assigned to the `e.Value` parameter of a `CalculateDocumentVariable` event handler.

## Files to Review

* [Form1.cs](./CS/Form1.cs) (VB: [Form1.vb](./VB/Form1.vb))
* [Program.cs](./CS/Program.cs) (VB: [Program.vb](./VB/Program.vb))

## More Examples

* [How to Use Document Variable (DOCVARIABLE) Fields to Insert Dynamic Content](https://github.com/DevExpress-Examples/how-to-use-docvariable-fields)
* [How to Implement Mail Merge in a RichEditControl](https://github.com/DevExpress-Examples/mail-merge-in-a-richeditcontrol)

## Documentation

* [DOCVARIABLE Field](https://docs.devexpress.com/WindowsForms/9721/controls-and-libraries/rich-text-editor/fields/field-codes/docvariable)
* [How to: Insert Dynamic Content](https://docs.devexpress.com/WindowsForms/401204/controls-and-libraries/rich-text-editor/examples/automation/how-to-insert-dynamic-content)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-richedit-use-calculatedocumentvariable-event-to-insert-formatted-content&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-richedit-use-calculatedocumentvariable-event-to-insert-formatted-content&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
