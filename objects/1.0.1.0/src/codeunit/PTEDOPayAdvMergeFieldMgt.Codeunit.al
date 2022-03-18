codeunit 62042 "PTE DO PayAdv MergeField Mgt"
{
    TableNo = "CDO E-Mail Codeunit Parameter";

    trigger OnRun()
    var
        RecRef: RecordRef;
    begin

        RecRef.OPEN(Rec."Table No.");
        RecRef.SETPOSITION(Rec.Position);
        IF NOT RecRef.FIND THEN
            EXIT;
        CASE Rec.Parameter OF
            'CURRENCYSYMBOL':
                Rec.ReturnValue := GetCurrencyCode(RecRef, true);
            'CURRENCYCODE':
                Rec.ReturnValue := GetCurrencyCode(RecRef, false);
        END;
    end;

    local procedure GetCurrencyCode(RecRef: RecordRef; ReturnCurrencySymbol: Boolean): Text
    var
        GenLedgerSetup: Record "General Ledger Setup";
        PaymentProposalHead: Record "OPP Payment Proposal Head";
        Currency: Record Currency;
    begin
        if RecRef.Number <> 5157896 then
            exit;

        RecRef.SetTable(PaymentProposalHead);
        if PaymentProposalHead."Pmt. Currency Code" = '' then begin
            GenLedgerSetup.Get();
            if ReturnCurrencySymbol then
                exit(GenLedgerSetup.GetCurrencySymbol())
            else
                exit(GenLedgerSetup.GetCurrencyCode(''));
        end else begin
            if not ReturnCurrencySymbol then
                exit(PaymentProposalHead."Pmt. Currency Code");

            if Currency.Get(PaymentProposalHead."Pmt. Currency Code") then
                exit(Currency.GetCurrencySymbol());
        end;
    end;

}
