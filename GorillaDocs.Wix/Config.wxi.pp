<?xml version="1.0" encoding="utf-8"?>
<Include>
  <?define Property_ProductVersion = "!(bind.FileVersion.[Enter assembly that MSI version should linked])" ?>
  <?define Property_AssemblyPath = "..\[Folder]\bin\$(var.Configuration)" ?>
  <!-- Platform variables -->
  <?if $(var.Platform) = x64 ?>
  <?define Property_ProductName = "[Enter Product Name]" ?>
  <?define Property_Win64 = "yes" ?>
  <?define Property_ProgramFilesFolder = "ProgramFiles64Folder" ?>
  <?define Property_CommonFilesFolder = "CommonFiles64Folder" ?>
  <?else ?>
  <?define Property_ProductName = "[Enter Product Name]" ?>
  <?define Property_Win64 = "no" ?>
  <?define Property_ProgramFilesFolder = "ProgramFilesFolder" ?>
  <?define Property_CommonFilesFolder = "CommonFilesFolder" ?>
  <?endif ?>
</Include>
