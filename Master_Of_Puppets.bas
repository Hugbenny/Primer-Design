Attribute VB_Name = "Master_Of_Puppets"
Sub Master_Of_Puppets()
    ' Call the first macro
    Call getDNA.getDNA_Coordinates
        
    ' Call the second macro
    Call Primer_Blast.Primer_Blast
    
    ' Call the third macro
    Call UCSC_PCR.UCSC_PCR
    
    ' Call the fourth macro
    Call SNP_Check.SNP_Check

End Sub
Sub Master_Of_Puppets_Gene()
    ' Call the first macro
    Call getDNA.getDNA_Gene
        
    ' Call the second macro
    Call Primer_Blast.Primer_Blast
    
    ' Call the third macro
    Call UCSC_PCR.UCSC_PCR
    
    ' Call the fourth macro
    Call SNP_Check.SNP_Check

End Sub
