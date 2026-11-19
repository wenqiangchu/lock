Attribute VB_Name = "TestSuite"
Option Explicit

Public Sub RunAllVbaTests()
    RunMdlInfraTests
    RunMdlSchemaTests
    RunMdlDiffTests
    RunMdlMergeTests
    RunMdlStorageTests
    RunMdlLockTests
    RunMdlSessionTests
    RunMdlExcelTests
    RunMdlUITests
End Sub
