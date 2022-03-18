report 62040 "PTE DO Payment Advice"
{

    RDLCLayout = 'src/report/PTEDOPaymentAdvice.Report.rdlc';

    Caption = 'Payment Advice';
    EnableHyperlinks = true;
    ApplicationArea = All;
    UsageCategory = ReportsAndAnalysis;

    dataset
    {
        dataitem(SingleRowData; "Integer")
        {
            DataItemTableView = SORTING(Number) Where(Number = const(1));
            column(CompanyAddr1; CompanyAddr[1])
            {
            }
            column(CompanyAddr2; CompanyAddr[2])
            {
            }
            column(CompanyAddr3; CompanyAddr[3])
            {
            }
            column(CompanyAddr4; CompanyAddr[4])
            {
            }
            column(CompanyAddr5; CompanyAddr[5])
            {
            }
            column(CompanyAddr6; CompanyAddr[6])
            {
            }
            column(PhoneNo_CompanyInfo; CompanyInfo."Phone No.")
            {
            }
            column(FaxNo_CompanyInfo; CompanyInfo."Fax No.")
            {
            }
            column(VATRegistrationNo_CompanyInfo; CompanyInfo."VAT Registration No.")
            {
            }
            column(Name_Address_CompanyInfo; CompanyInfo.Name + ' ' + CompanyInfo.Address + ' ' + CompanyInfo."Post Code" + ' ' + CompanyInfo.City)
            {
            }
            column(HyperlinkColor; PmtExpSetup."Hyperlink Color")
            {
            }
            column(DocumentNoType; Format(DocumentNoType, 0, 2))
            {
            }
        }
        dataitem("OPP Payment Proposal"; "OPP Payment Proposal")
        {
            DataItemTableView = SORTING("Journal Template Name", "Journal Batch Name");
            PrintOnlyIfDetail = true;
            RequestFilterFields = "Journal Template Name", "Journal Batch Name";
            column(JnlTempName_PaymentProposal; "Journal Template Name")
            {
            }
            column(JnlBatchName_PaymentProposal; "Journal Batch Name")
            {
            }
            dataitem("OPP Payment Proposal Head"; "OPP Payment Proposal Head")
            {
                DataItemLink = "Gen. Journal Template" = field("Journal Template Name"), "Gen. Journal Batch" = field("Journal Batch Name");
                DataItemTableView = SORTING("Gen. Journal Template", "Gen. Journal Batch", "Gen. Journal Line");
                RequestFilterFields = "Gen. Journal Line", "Account No.", "Print Payment Advice";
                column(GenJnlTemplate_PaymentHeader; "Gen. Journal Template")
                {
                }
                column(GenJnlBatch_PaymentHeader; "Gen. Journal Batch")
                {
                }
                column(GenJnlLine_PaymentHeader; "Gen. Journal Line")
                {
                }
                dataitem(PageLoop; "Integer")
                {
                    DataItemTableView = SORTING(Number) Where(Number = const(1));
                    dataitem(HeaderData; "Integer")
                    {
                        DataItemTableView = SORTING(Number) Where(Number = const(1));
                        column(PostingDate_PaymentProposal; Format("OPP Payment Proposal"."Posting Date"))
                        {
                        }
                        column(FaxNo_PaymentHeader; "OPP Payment Proposal Head"."Fax No.")
                        {
                        }
                        column(EMailNo_PaymentHeader; "OPP Payment Proposal Head"."E-Mail")
                        {
                        }
                        column(OurAccountNo_PaymentHeader; "OPP Payment Proposal Head"."Our Account No.")
                        {
                        }
                        column(OrdererBankName_PaymentHeader; "OPP Payment Proposal Head"."Orderer Bank Name")
                        {
                        }
                        column(PaymentTypeCode_PaymentHeader; "OPP Payment Proposal Head"."Payment Type Code")
                        {
                        }
                        column(PmtAddr1; PmtAddr[1])
                        {
                        }
                        column(PmtAddr2; PmtAddr[2])
                        {
                        }
                        column(PmtAddr3; PmtAddr[3])
                        {
                        }
                        column(PmtAddr4; PmtAddr[4])
                        {
                        }
                        column(PmtAddr5; PmtAddr[5])
                        {
                        }
                        column(PmtAddr6; PmtAddr[6])
                        {
                        }
                        column(PmtAddr7; PmtAddr[7])
                        {
                        }
                        column(PmtAddr8; PmtAddr[8])
                        {
                        }
                        column(AccountName; AccountName)
                        {
                        }
                        column(MandateID; TextMandateID)
                        {
                        }
                        column(CredID; TextCredID)
                        {
                        }
                        column(BankName; BankName)
                        {
                        }
                        column(PageGroupNo; PageGroupNo)
                        {
                        }
                        column(NextPageGroupNo; NextPageGroupNo)
                        {
                        }
                        column(PageLoop_Number; PageLoop.Number)
                        {
                        }
                        column(Text006; Text006)
                        {
                        }
                        column(Text1010; Text1010)
                        {
                        }
                        column(Text1011; Text1011)
                        {
                        }
                        column(OurAccountNo_PaymentHeaderCaption; Payment_Proposal_Head___Our_Account_No__CaptionLbl)
                        {
                        }
                        column(OrdererBankName_PaymentHeaderCaption; Payment_Proposal_Head___Orderer_Bank_Name_CaptionLbl)
                        {
                        }
                        column(PhoneNo_CompanyInfoCaption; CompanyInfo__Phone_No__CaptionLbl)
                        {
                        }
                        column(FaxNo_CompanyInfoCaption; CompanyInfo__Fax_No__CaptionLbl)
                        {
                        }
                        column(VATRegistrationNo_CompanyInfoCaption; CompanyInfo__VAT_Registration_No__CaptionLbl)
                        {
                        }
                        column(GesamtCaption; StrSubStNo(Text011, CurrCode))
                        {
                        }
                        column(DocTypeCaption; DocTypeLbl)
                        {
                        }
                        column(DocNoCaption; DocNoLbl)
                        {
                        }
                        column(AmountCaption; AmountLbl)
                        {
                        }
                        column(PmtDiscountCaption; PmtDiscountLbl)
                        {
                        }
                        column(PmtAmountCaption; PmtAmountLbl)
                        {
                        }
                        column(PaymentInformationCaption; PaymentInformationLbl)
                        {
                        }
                        column(ExtDocNoCaption; ExtDocNoLbl)
                        {
                        }
                        column(URL_Mandate; URL_Mandate)
                        {
                        }
                        column(URL_Account; URL_Account)
                        {
                        }


                        trigger OnAfterGetRecord()
                        begin
                            // RecordRef for controlling addresses from the report (Go to URL) for customer and vendor
                            Clear(URL_Account);
                            if "OPP Payment Proposal Head"."Account No." <> '' then
                                if "OPP Payment Proposal Head"."Account Type" = "OPP Payment Proposal Head"."Account Type"::Customer then begin
                                    Cust.Get("OPP Payment Proposal Head"."Account No.");
                                    RecRef.GetTable(Cust);
                                    URL_Account := GenerateHyperlink(Format(RecRef.RecordID, 0, 10), Page::"Customer Card");
                                end else
                                    if "OPP Payment Proposal Head"."Account Type" = "OPP Payment Proposal Head"."Account Type"::Vendor then begin
                                        Vend.Get("OPP Payment Proposal Head"."Account No.");
                                        RecRef.GetTable(Vend);
                                        URL_Account := GenerateHyperlink(Format(RecRef.RecordID, 0, 10), Page::"Vendor Card");
                                    end;

                            // RecordRef URL adress for mandate
                            Clear(URL_Mandate);
                            if "OPP Payment Proposal Head"."Mandate ID" <> '' then begin
                                BankAccountMandate.Get("OPP Payment Proposal Head"."Mandate ID");
                                RecRef.GetTable(BankAccountMandate);
                                URL_Mandate := GenerateHyperlink(Format(RecRef.RecordID, 0, 10), Page::"OPP Bank Account Mandate");
                            end;


                        end;
                    }
                    dataitem("OPP Payment Proposal Line"; "OPP Payment Proposal Line")
                    {
                        DataItemLink = "Journal Template Name" = field("Gen. Journal Template"), "Journal Batch Name" = field("Gen. Journal Batch"), "Journal Line No." = field("Gen. Journal Line");
                        DataItemLinkReference = "OPP Payment Proposal Head";
                        DataItemTableView = SORTING("Journal Template Name", "Journal Batch Name", "Journal Line No.", "Line No.");
                        column(LineNo_PaymentProposalLine; "Line No.")
                        {
                        }
                        column(IDAppliedEntry_PaymentProposalLine; "ID Applied-Entry")
                        {
                        }
                        column(OriginalRemainingAmount_PaymentLine; "Original Remaining Amount")
                        {
                        }
                        column(PostingPaymentDiscount_PaymentLine; "Posting Payment Discount")
                        {
                        }
                        column(PostingAppliedAmount_PaymentLine; "Posting Applied Amount")
                        {
                        }
                        column(AppliesToDocType_PaymentLine; CopyStr(Format("Applies-to Doc. Type"), 1, 2))
                        {
                        }
                        column(AppliesToDocNo_PaymentLine; "Applies-to Doc. No.")
                        {
                        }
                        column(PaymentText_PaymentLine; "Payment Text")
                        {
                        }
                        column(ExternalDocumentNo_PaymentLine; "External Document No.")
                        {
                        }
                        column(OriginalRemainingAmount_PaymentLine_Total; hSum[1])
                        {
                        }
                        column(PostingPaymentDiscount_PaymentLine_Total; hSum[2])
                        {
                        }
                        column(PostingAppliedAmount_PaymentLine_Total; hSum[3])
                        {
                        }
                        dataitem("OPP Ledger Entry Comment Line"; "OPP Ledger Entry Comment Line")
                        {
                            DataItemLink = "No." = field("Account No."), "Entry No." = field("ID Applied-Entry");
                            DataItemTableView = SORTING("Table Name", "No.", "Entry No.", "Line No.");
                            column(Comment_CommentLine; Comment)
                            {
                            }
                            column(LineNo_CommentLine; "Line No.")
                            {
                            }

                            trigger OnPreDataItem()
                            begin
                                if not IncludeComments then
                                    CurrReport.Break();

                                if "OPP Payment Proposal Line"."Account Type" = "OPP Payment Proposal Line"."Account Type"::Vendor then
                                    SetRange("Table Name", "Table Name"::Vendor)
                                else
                                    if "OPP Payment Proposal Line"."Account Type" = "OPP Payment Proposal Line"."Account Type"::Customer then
                                        SetRange("Table Name", "Table Name"::Customer)
                                    else
                                        if "OPP Payment Proposal Line"."Account Type" = "OPP Payment Proposal Line"."Account Type"::"G/L Account" then
                                            SetRange("Table Name", "Table Name"::"G/L Account");

                                if not FindFirst() then
                                    CurrReport.Break();
                            end;
                        }

                        trigger OnAfterGetRecord()
                        begin
                            if ("Posting Payment Discount" = 0) and ("Posting Applied Amount" = 0) then
                                CurrReport.Skip();

                            if PaymentType."Payment Direction" = PaymentType."Payment Direction"::Outgoing then begin
                                "Original Remaining Amount" := -"Original Remaining Amount";
                                "Posting Payment Discount" := -"Posting Payment Discount";
                                "Posting Applied Amount" := -"Posting Applied Amount";
                            end;

                            hSum[1] += "Original Remaining Amount";
                            hSum[2] += "Posting Payment Discount";
                            hSum[3] += "Posting Applied Amount";

                            if ExcelOutput then begin
                                TempOPPPmtProposalLineExcel := "OPP Payment Proposal Line";
                                TempOPPPmtProposalLineExcel.Insert();
                            end;

                        end;

                    }

                    trigger OnPreDataItem()
                    begin
                        Clear(hSum);
                    end;
                }

                trigger OnAfterGetRecord()
                var
                    ExecutionDate: Date;
                    CountryRegionCode: Code[10];
                    TargetDueDays: Integer;
                begin
                    PageGroupNo += 1;
                    if Language.GetLanguageID("Language Code") <> 0 then
                        CurrReport.Language := Language.GetLanguageID("Language Code");
                    if not PaymentType.Get("Payment Type Code") then
                        PaymentType.Init();

                    CountryRegionCode := PmtExpSetup.GetPmtCountryReverse("Country/Region Code", false);
                    FormatAddr.FormatAddr(PmtAddr, Name, "Name 2", '', Address, "Address 2", City, "Post Code", '', CountryRegionCode);

                    if PaymentType."Bank Branch Code required" then
                        BankName := CopyStr("Bank Name" + ' ' + Text012 + ': ' + "Bank Branch No." + ' ' + Text013 + ': ' + "Bank Account No.",
                            1, MaxStrLen(BankName))
                    else begin
                        BankName := "Bank Name";
                        if "Payment Type Code" = OPPClearingISO.GetISOPMT() then
                            if "Routing No." <> '' then
                                BankName := CopyStr(BankName + ' ' + FieldCaption("Routing No.") + ': ' + "Routing No.", 1, MaxStrLen(BankName));

                        if "Bank BIC Code" <> '' then
                            BankName := CopyStr(BankName + ' BIC: ' + "Bank BIC Code", 1, MaxStrLen(BankName))
                        else
                            if "Bank Branch No." <> '' then
                                BankName := CopyStr(BankName + ' ' + Text012 + ': ' + "Bank Branch No.", 1, MaxStrLen(BankName));

                        if "Bank IBAN Code" <> '' then
                            BankName := CopyStr(BankName + ' IBAN: ' + OPPPmtExpTools.PrintIBANInBlocks("Bank IBAN Code"), 1, MaxStrLen(BankName))
                        else
                            BankName := CopyStr(BankName + ' ' + Text013 + ': ' + "Bank Account No.", 1, MaxStrLen(BankName));
                    end;
                    AccountName := Format("Account Type") + ' ' + "Account No.";
                    if not PrintFaxNo then
                        "Fax No." := '';
                    if not PrintEMail then
                        "E-Mail" := '';

                    CredID := PmtExpSetup."Creditor Identifier";
                    MandateID := "Mandate ID";
                    if "Account Type" = "Account Type"::Customer then begin
                        if CustBankAccount.Get("Account No.", "Bank Code") then
                            if CustBankAccount."OPP Alt. Creditor Identifier" <> '' then
                                CredID := CustBankAccount."OPP Alt. Creditor Identifier";
                    end else
                        if "Account Type" = "Account Type"::Vendor then
                            if VendBankAccount.Get("Account No.", "Bank Code") then
                                if VendBankAccount."OPP Alt. Creditor Identifier" <> '' then
                                    CredID := VendBankAccount."OPP Alt. Creditor Identifier";

                    if OPPPmtExpTools.IsSepaDDPmtType("Payment Type Code") then begin
                        TextCredID := CopyStr(PmtExpSetup.FieldCaption("Creditor Identifier") + ': ' + CredID, 1, MaxStrLen(TextCredID));
                        TextMandateID := CopyStr(FieldCaption("Mandate ID") + ': ' + "OPP Payment Proposal Head"."Mandate ID", 1, MaxStrLen(TextMandateID));
                    end else begin
                        TextCredID := '';
                        TextMandateID := '';
                    end;

                    ExecutionDate := "OPP Payment Proposal"."Execution Date";
                    if ExecutionDate = 0D then
                        ExecutionDate := "OPP Payment Proposal"."Posting Date";
                    if PmtExpSetup."SEPA Due in Pmt. File" then
                        if "SEPA Due Days" > 0 then
                            OPPPmtExpTools.CalcSEPATargetDueDate(ExecutionDate, "SEPA Due Days", TargetDueDays, ExecutionDate);
                    if PaymentType."Payment Direction" = PaymentType."Payment Direction"::Incoming then begin
                        Text1010 := Text002;
                        Text1011 := StrSubStNo(Text004, ExecutionDate);
                    end else begin
                        Text1010 := Text001;
                        Text1011 := StrSubStNo(Text003, ExecutionDate);
                    end;
                    if "Pmt. Currency Code" <> '' then
                        CurrCode := "Pmt. Currency Code"
                    else
                        CurrCode := GeneralLedgerSetup."LCY Code";

                    if ExcelOutput then begin
                        TempOPPPmtProposalHeadExcel := "OPP Payment Proposal Head";
                        TempOPPPmtProposalHeadExcel.Insert();
                    end;
                end;

                trigger OnPreDataItem()
                begin
                    if (GetFilter("Gen. Journal Template") = '') and ("OPP Payment Proposal".GetFilter("Journal Template Name") <> '') then
                        SetRange("Gen. Journal Template", "OPP Payment Proposal".GetFilter("Journal Template Name"));

                    if (GetFilter("Gen. Journal Batch") = '') and ("OPP Payment Proposal".GetFilter("Journal Batch Name") <> '') then
                        SetRange("Gen. Journal Batch", "OPP Payment Proposal".GetFilter("Journal Batch Name"));
                end;
            }

            trigger OnPreDataItem()
            begin
                PageGroupNo := 1;
                NextPageGroupNo := 1;

                if (GetFilter("Journal Template Name") = '') and ("OPP Payment Proposal Head".GetFilter("Gen. Journal Template") <> '') then
                    SetRange("Journal Template Name", "OPP Payment Proposal Head".GetFilter("Gen. Journal Template"));

                if (GetFilter("Journal Batch Name") = '') and ("OPP Payment Proposal Head".GetFilter("Gen. Journal Batch") <> '') then
                    SetRange("Journal Batch Name", "OPP Payment Proposal Head".GetFilter("Gen. Journal Batch"));
            end;
        }
    }

    RequestPage
    {
        Caption = 'Payment Advice';
        SaveValues = true;

        layout
        {
            area(content)
            {
                group(Control5157805)
                {
                    ShowCaption = false;
                    field(PrintFaxNo; PrintFaxNo)
                    {
                        ApplicationArea = All;
                        Caption = 'Print Fax No.';
                        ToolTip = 'By selecting the Print Fax No. option, the fax number of the customer/vendor specified in the payment proposal header will be printed in the top left corner of the report.';
                    }
                    field(PrintEMail; PrintEMail)
                    {
                        ApplicationArea = All;
                        Caption = 'Print E-Mail';
                        ToolTip = 'The Print e-mail option prints the customer/vendor e-mail address in the payment proposal header on the payment advice note.';
                    }
                    field(ExcelOutput; ExcelOutput)
                    {
                        ApplicationArea = All;
                        Caption = 'Output in Excel';
                        ToolTip = 'A check mark in this field indicates that the report contents are transferred to an Excelsheet.';
                    }
                    field(IncludeComments; IncludeComments)
                    {
                        ApplicationArea = All;
                        Caption = 'Incl. Comments';
                        ToolTip = 'Specifies whether any existing entry comments should be printed below the respective entry.';
                    }
                    field(DocumentNoType; DocumentNoType)
                    {
                        ApplicationArea = All;
                        Caption = 'Show Document No.';
                        OptionCaption = 'Document No. and External Document No.,External Document No.,Document No.';
                        ToolTip = 'Specifies which document no. field should be printed.';
                    }
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPreReport()
    begin

        CompanyInfo.Get();
        PmtExpSetup.Get();
        GeneralLedgerSetup.Get();

        FormatAddr.Company(CompanyAddr, CompanyInfo);
    end;

    trigger OnPostReport()
    begin
        if ExcelOutput then begin
            Commit();
            ExportExcelValuesAfterPostReport();
        end;
    end;


    var
        PmtExpSetup: Record "OPP Payment Export Setup";
        CompanyInfo: Record "Company Information";
        PaymentType: Record "OPP Payment Type Code";
        GeneralLedgerSetup: Record "General Ledger Setup";
        CustBankAccount: Record "Customer Bank Account";
        VendBankAccount: Record "Vendor Bank Account";
        BankAccountMandate: Record "OPP Bank Account Mandate";
        Cust: Record Customer;
        Vend: Record Vendor;
        TempOPPPmtProposalLineExcel: Record "OPP Payment Proposal Line" temporary;
        TempOPPPmtProposalHeadExcel: Record "OPP Payment Proposal Head" temporary;
        Language: Codeunit Language;
        FormatAddr: Codeunit "Format Address";
        OPPPmtExpTools: Codeunit "OPP Pmt. Export Tools";
        OPPTriggerMgt: Codeunit "OPP Trigger Management";
        ExcelTempBlob: Codeunit "Temp Blob";
        OPPClearingISO: Codeunit "OPP Clearing ISO";
        RecRef: RecordRef;
        ExcelStream: OutStream;
        CurrCode: Code[10];
        CompanyAddr: Array[8] of Text[100];
        PmtAddr: Array[8] of Text[100];
        BankName: Text[250];
        AccountName: Text[250];
        Line: Text[1024];
        Text1010: Text[100];
        Text1011: Text[100];
        TextCredID: Text[80];
        TextMandateID: Text[80];
        ExcelFileName: Text[1024];
        [InDataSet]
        CredID: Text[250];
        [InDataSet]
        MandateID: Text[250];
        URL_Mandate: Text;
        URL_Account: Text;
        hSum: Array[3] of Decimal;
        PageGroupNo: Integer;
        NextPageGroupNo: Integer;
        DocumentNoType: Option "0","1","2";
        [InDataSet]
        PrintEMail: Boolean;
        [InDataSet]
        ExcelOutput: Boolean;
        [InDataSet]
        PrintFaxNo: Boolean;
        WriteLineNo: Boolean;
        [InDataSet]
        IncludeComments: Boolean;
        Text001: Label 'Payment Advice';
        Text002: Label 'Direct Debit Advice';
        Text003: Label 'We will pay the following amounts per %1 to your bank account shown below:';
        Text004: Label 'We will charge the following amounts per %1 to your account shown below:';
        Text006: Label 'Page:';
        Text011: Label 'Total in %1';
        CompanyInfo__Phone_No__CaptionLbl: Label 'Phone No.';
        CompanyInfo__Fax_No__CaptionLbl: Label 'Fax No.';
        CompanyInfo__VAT_Registration_No__CaptionLbl: Label 'VAT Reg. No.';
        Payment_Proposal_Head___Our_Account_No__CaptionLbl: Label 'Our Account No.';
        Payment_Proposal_Head___Orderer_Bank_Name_CaptionLbl: Label 'Our Bank';
        DocTypeLbl: Label 'Doc. Type';
        DocNoLbl: Label 'Doc. No.';
        AmountLbl: Label 'Amount';
        PmtDiscountLbl: Label 'Pmt. Discount';
        PmtAmountLbl: Label 'Pmt. Amount';
        PaymentInformationLbl: Label 'Payment Information';
        ExtDocNoLbl: Label 'Ext. Doc. No.';
        Text012: Label 'Branch No.';
        Text013: Label 'Bank Account No.';

    procedure SetData(PDF_PaymentProposal: Record "OPP Payment Proposal"; PDF_PaymentProposalHead: Record "OPP Payment Proposal Head"; PDF_FaxNr: Boolean)
    begin
        "OPP Payment Proposal" := PDF_PaymentProposal;
        "OPP Payment Proposal Head" := PDF_PaymentProposalHead;
        PrintFaxNo := PDF_FaxNr;
    end;

    procedure SetFaxOutput(DoFAX: Boolean)
    begin
        PrintFaxNo := DoFAX;
    end;

    local procedure ExportExcelValuesAfterPostReport()
    var
        OPPPmtProposalExcel: Record "OPP Payment Proposal";
        PaymentProposalHead2: Record "OPP Payment Proposal Head";
        TempBlobExcel: Codeunit "Temp Blob";
        StreamExcel: OutStream;
    begin
        if TempOPPPmtProposalHeadExcel.FindSet() then
            repeat
                Clear(ExcelStream);
                Clear(ExcelTempBlob);
                ExcelTempBlob.CreateOutStream(ExcelStream, TextEncoding::UTF8);

                PaymentProposalHead2.SetRange("Gen. Journal Template", TempOPPPmtProposalHeadExcel."Gen. Journal Template");
                PaymentProposalHead2.SetRange("Gen. Journal Batch", TempOPPPmtProposalHeadExcel."Gen. Journal Batch");
                PaymentProposalHead2.SetRange("Account Type", TempOPPPmtProposalHeadExcel."Account Type");
                PaymentProposalHead2.SetRange("Account No.", TempOPPPmtProposalHeadExcel."Account No.");
                WriteLineNo := PaymentProposalHead2.Count > 1;
                if WriteLineNo then
                    ExcelFileName := CopyStr(Format(TempOPPPmtProposalHeadExcel."Account Type"), 1, 1) +
                      TempOPPPmtProposalHeadExcel."Account No." + '_' + Format(TempOPPPmtProposalHeadExcel."Gen. Journal Line") + '.CSV'
                else
                    ExcelFileName := CopyStr(Format(TempOPPPmtProposalHeadExcel."Account Type"), 1, 1) +
                      TempOPPPmtProposalHeadExcel."Account No." + '.CSV';

                Line :=
                  "OPP Payment Proposal Line".FieldCaption("Pmt. Account Type") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Pmt. Account No.") + ';' +
                  "OPP Payment Proposal Head".FieldCaption(Name) + ';' +
                  "OPP Payment Proposal".FieldCaption("Document No.") + ';' +
                  "OPP Payment Proposal".FieldCaption("Posting Date") + ';' +
                  "OPP Payment Proposal".FieldCaption("Bank Branch Code") + ';' +
                  "OPP Payment Proposal".FieldCaption("Bank Account No.") + ';' +
                  "OPP Payment Proposal".FieldCaption("Bank BIC Code") + ';' +
                  "OPP Payment Proposal".FieldCaption("Bank IBAN") + ';' +
                  "OPP Payment Proposal".FieldCaption("Document No.") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Account Type") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Account No.") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Association Sort Code") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Orig. Posting Date") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Document Date") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Due Date") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Discount Date") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Applies-to Doc. Type") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Applies-to Doc. No.") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("External Document No.") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Original Amount") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Orig. Currency Code") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Original Amt. (LCY)") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Original Remaining Amount") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Appln. Remaining Amount") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Payment Description") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Pmt. Currency Code") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Posting Applied Amount") + ';' +
                  "OPP Payment Proposal Line".FieldCaption("Posting Payment Discount");
                ExcelStream.WriteText(Line);
                ExcelStream.WriteText();
                TempOPPPmtProposalLineExcel.SetRange("Journal Template Name", TempOPPPmtProposalHeadExcel."Gen. Journal Template");
                TempOPPPmtProposalLineExcel.SetRange("Journal Batch Name", TempOPPPmtProposalHeadExcel."Gen. Journal Batch");
                TempOPPPmtProposalLineExcel.SetRange("Journal Line No.", TempOPPPmtProposalHeadExcel."Gen. Journal Line");

                if TempOPPPmtProposalLineExcel.FindSet() then
                    repeat
                        OPPPmtProposalExcel.Get(TempOPPPmtProposalLineExcel."Journal Template Name", TempOPPPmtProposalLineExcel."Journal Batch Name");

                        Line :=
                          Format(TempOPPPmtProposalLineExcel."Pmt. Account Type") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Pmt. Account No.") + ';' +
                          Format(TempOPPPmtProposalHeadExcel.Name) + ';' +
                          Format(OPPPmtProposalExcel."Document No.") + ';' +
                          Format(OPPPmtProposalExcel."Posting Date") + ';' +
                          Format(OPPPmtProposalExcel."Bank Branch Code") + ';' +
                          Format(OPPPmtProposalExcel."Bank Account No.") + ';' +
                          Format(OPPPmtProposalExcel."Bank BIC Code") + ';' +
                          Format(OPPPmtProposalExcel."Bank IBAN") + ';' +
                          Format(OPPPmtProposalExcel."Document No.") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Account Type") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Account No.") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Association Sort Code") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Orig. Posting Date") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Document Date") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Due Date") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Discount Date") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Applies-to Doc. Type") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Applies-to Doc. No.") + ';' +
                          Format(TempOPPPmtProposalLineExcel."External Document No.") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Original Amount") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Orig. Currency Code") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Original Amt. (LCY)") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Original Remaining Amount") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Appln. Remaining Amount") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Payment Description") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Pmt. Currency Code") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Posting Applied Amount") + ';' +
                          Format(TempOPPPmtProposalLineExcel."Posting Payment Discount");
                        ExcelStream.WriteText(Line);
                        ExcelStream.WriteText();
                    until TempOPPPmtProposalLineExcel.Next() = 0;
                OPPTriggerMgt.BLOBExportGateway(ExcelTempBlob, PmtExpSetup."Pmt. Advice File Path" + ExcelFileName, '');
                Commit();
            until TempOPPPmtProposalHeadExcel.Next() = 0;
    end;

    procedure GenerateHyperlink(Bookmark: Text[250]; PageID: Integer): Text
    var
        ClientTypeManagement: Codeunit "Client Type Management";
    begin
        if Bookmark = '' then
            exit('');
        // Generates a URL such as dynamicsnav://hostname:port/instance/company/runpage?page=pageId<&Tenant=tenantId>&bookmark=recordId&mode=View.
        exit(GetUrl(ClientType::Web, CompanyName, ObjectType::Page, PageID) +
            StrSubStNo('&mode=view&bookmark=%1', Bookmark));
        //STRSUBSTNO('&amp;bookmark=%1&mode=View',Bookmark));
    end;

}

