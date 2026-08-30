Attribute VB_Name = "lettretype"

Public Sub MAIN()
Dim x
Dim y
WordBasic.BeginDialog 308, 258, "Lettres Types"
    WordBasic.OKButton 10, 6, 259, 41
    WordBasic.CancelButton 174, 213, 88, 21
    WordBasic.OptionGroup "GroupeOpts6"
        WordBasic.OptionButton 10, 52, 205, 16, "Lettre Type Word", "BoutonOpt214"
        WordBasic.OptionButton 10, 68, 180, 17, "Lettre Type DBmed", "BoutonOpt215"
        WordBasic.OptionButton 10, 86, 148, 16, "Lettre vierge", "BoutonOpt216"
        WordBasic.OptionButton 10, 103, 148, 16, "Présence", "BoutonOpt217"
        WordBasic.OptionButton 10, 120, 148, 16, "Bouton d'option", "BoutonOpt218"
        WordBasic.OptionButton 10, 137, 148, 16, "Bouton d'option", "BoutonOpt219"
        WordBasic.OptionButton 10, 154, 148, 16, "Bouton d'option", "BoutonOpt220"
        WordBasic.OptionButton 10, 171, 148, 16, "Bouton d'option", "BoutonOpt221"
        WordBasic.OptionButton 10, 188, 148, 16, "Bouton d'option", "BoutonOpt222"
        WordBasic.OptionButton 10, 205, 148, 16, "Bouton d'option", "BoutonOpt223"
        WordBasic.OptionButton 10, 222, 148, 16, "Bouton d'option", "BoutonOpt224"
WordBasic.EndDialog
Dim dlg As Object: Set dlg = WordBasic.CurValues.UserDialog
x = WordBasic.Dialog.UserDialog(dlg, 1)
If x = 0 Then GoTo fin
y = dlg.GroupeOpts6
If y = 0 Then WordBasic.Call "LTW1"
If y = 1 Then WordBasic.Call "DBmed1"
If y = 2 Then WordBasic.Call "LettreV"
If y = 3 Then WordBasic.Call "Presence"
If y = 4 Then WordBasic.Call "MedecinsE"
If y = 5 Then WordBasic.Call "MedecinsF"
fin:
End Sub
