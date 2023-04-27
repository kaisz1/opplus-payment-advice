codeunit 62040 "PTE DO Payment Advice Mgt."
{
    procedure SendAdvice(PaymentProposal: Record "OPP Payment Proposal"): Boolean
    var
        EMailTemplateLine: Record "CDO E-Mail Template Line";
        PaymentProposalHeads: Record "OPP Payment Proposal Head";
        PaymentProposalHead: Record "OPP Payment Proposal Head";
        PaymentExportSetup: Record "OPP Payment Export Setup";
        CustomReportSelection: Record "Custom Report Selection";
        FilterRecord: RecordRef;
        VariantRecord: Variant;
        QueueCounter: Integer;
        PaymentAvisReportID: Integer;
        PaymentProposalsQueuedLbl: Label '%1 payment proposal(s) have been queued to be send by Document Output.', Comment = ' %1 = quantity of queued payment proposal heads';
    begin
        if not PaymentExportSetup.Get() then
            exit;

        // Iterate through the paym. proposal heads and find the once that should be send/printed
        PaymentProposalHeads.SetRange("Gen. Journal Template", PaymentProposal."Journal Template Name");
        PaymentProposalHeads.SetRange("Gen. Journal Batch", PaymentProposal."Journal Batch Name");
        PaymentProposalHeads.SetRange("Print Payment Advice", true);

        if PaymentProposalHeads.FindSet() then
            repeat
                PaymentProposalHead := PaymentProposalHeads;
                PaymentProposalHead.SetRecFilter();

                // Try to find template based on the report id that is configured in opp paym. setup and the current language code
                FilterRecord.GETTABLE(PaymentProposalHead);
                FilterRecord.SetView(PaymentProposalHead.GetView());

                // Find ReportID in 'Customer Report Selection'
                if PaymentProposalHead."account type" = PaymentProposalHead."Account Type"::Customer then
                    CustomReportSelection.SetRange("Source Type", 18);
                if PaymentProposalHead."account type" = PaymentProposalHead."Account Type"::Vendor then
                    CustomReportSelection.SetRange("Source Type", 23);
                CustomReportSelection.SetRange("Source No.", PaymentProposalHead."Account No.");
                CustomReportSelection.SetRange(Usage, CustomReportSelection.Usage::"OPP-Advice");
                if CustomReportSelection.FindFirst() then
                    PaymentAvisReportID := CustomReportSelection."Report ID";

                if EMailTemplateLine.FindTemplate(PaymentAvisReportID, PaymentProposalHead."Language Code", FilterRecord) then begin
                    VariantRecord := PaymentProposalHead;
                    EMailTemplateLine.QueueMail(FilterRecord, VariantRecord, 0, 0);
                    QueueCounter += 1;
                end;

            until PaymentProposalHeads.Next() = 0;

        // Feedback to user about quantity of queued entries
        Message(PaymentProposalsQueuedLbl, QueueCounter);
    end;
}