codeunit 62040 "PTE DO Payment Advice Mgt."
{
    trigger OnRun()
    begin

    end;

    procedure SendAdvice(PaymentProposal: Record "OPP Payment Proposal"): Boolean
    var
        EMailTemplateLine: Record "CDO E-Mail Template Line";
        PaymentProposalHead: Record "OPP Payment Proposal Head";
        PaymentExportSetup: Record "OPP Payment Export Setup";
        FilterRecord: RecordRef;
        VariantRecord: Variant;
        QueueCounter: Integer;
        PaymentProposalsQueuedLbl: Label '%1 payment proposal(s) have been queued to be send by Document Output.', Comment = ' %1 = quantity of queued payment proposal heads';
    begin
        if not PaymentExportSetup.Get() then
            exit;

        // Iterate through the paym. proposal heads and find the once that should be send/printed
        PaymentProposalHead.SetRange("Gen. Journal Template", PaymentProposal."Journal Template Name");
        PaymentProposalHead.SetRange("Gen. Journal Batch", PaymentProposal."Journal Batch Name");
        PaymentProposalHead.SetRange("Print Payment Advice", true);

        if PaymentProposalHead.FindSet() then
            repeat
                // Try to find template based on the report id that is configured in opp paym. setup and the current language code
                FilterRecord.GETTABLE(PaymentProposalHead);
                FilterRecord.SetView(PaymentProposalHead.GetView());

                if EMailTemplateLine.FindTemplate(PaymentExportSetup."Payment Avis Report ID", PaymentProposalHead."Language Code", FilterRecord) then begin
                    VariantRecord := PaymentProposalHead;
                    EMailTemplateLine.QueueMail(FilterRecord, VariantRecord, 0, 0);
                    QueueCounter += 1;
                end;
            //end
            until PaymentProposalHead.Next() = 0;

        // Feedback to user about quantity of queued entries
        Message(PaymentProposalsQueuedLbl, QueueCounter);
    end;
}