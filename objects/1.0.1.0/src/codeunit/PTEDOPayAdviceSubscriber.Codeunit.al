codeunit 62041 "PTE DO Pay. Advice Subscriber"
{
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"CDO Events", 'OnFindTemplateWithReportID', '', true, true)]
    local procedure CDO_OnFindTemplateWithReportID(LanguageCode: Code[10]; ReportID: Integer; var EMailTemplateLine: Record "CDO E-Mail Template Line"; var RecRef: RecordRef; var TemplateFound: Boolean)
    var
        OppPaymentProposalHead: Record "OPP Payment Proposal Head";
        CdoEmailTemplateHeader: Record "CDO E-Mail Template Header";
    begin
        // check if recref is OPP Paym. Proposal head
        if RecRef.Number <> 5157896 then
            exit;

        RecRef.SetTable(OppPaymentProposalHead);

        case OppPaymentProposalHead."Account Type" of
            OppPaymentProposalHead."Account Type"::Customer:
                CdoEmailTemplateHeader.SetRange("Linked to", CdoEmailTemplateHeader."Linked to"::Customer);
            OppPaymentProposalHead."Account Type"::Vendor:
                CdoEmailTemplateHeader.SetRange("Linked to", CdoEmailTemplateHeader."Linked to"::Vendor);
            else
                exit;
        end;

        CdoEmailTemplateHeader.SetRange("Report-ID", ReportID);
        CdoEmailTemplateHeader.SetRange("First Table in Report", RecRef.Number);
        if CdoEmailTemplateHeader.IsEmpty then
            exit;

        CdoEmailTemplateHeader.FindFirst();
        repeat
            EMailTemplateLine.SetRange("E-Mail Template Code", CdoEmailTemplateHeader.Code);
            EMailTemplateLine.SetRange("Language Code", LanguageCode);
            EMailTemplateLine.SetRange(Enabled, true);
            TemplateFound := EMailTemplateLine.FindFirst();
            if not TemplateFound then begin
                EMailTemplateLine.SetRange("Language Code");
                TemplateFound := EMailTemplateLine.FindFirst();
            end;
        until (CdoEmailTemplateHeader.Next() = 0) or (TemplateFound);
    end;
}