Attribute VB_Name = "Mod_Devoir1v2"
Function Eq1IsoleXGenerale(y, a1, b1, c1)
' Renvoyer la valeur de x quand il est isolé dans la première équation
    x = (c1 - (b1 * y)) / a1
    
    Eq1IsoleXGenerale = x
    
End Function


Function Eq1IsoleYGenerale(x, a2, b2, c2)
' Renvoyer la valeur de y quand il est isolé dans la deuxième équation
    y = (c2 - (a2 * x)) / b2
    
    Eq1IsoleYGenerale = y
    
End Function


Function ResoudreGenerale(y0, n, a1, b1, c1, a2, b2, c2)
' Assigner la valeur de y0 à la variable y, qui pourra être utilisé dans le premier appel de Eq1IsoleXGenerale
    y = y0
    
    If n >= 1 Then
' Retourner à l'utilisateur les dernières valeurs de x et y obtenues après le nombre d'itérations voulues
        For i = 1 To n
        
            x = Eq1IsoleXGenerale(y, a1, b1, c1)
        
            y = Eq1IsoleYGenerale(x, a2, b2, c2)
        
        Next i
' Définir un vecteur de retour pour les résultats de x et y obtenus après n itérations
        Dim vecteur(1 To 2)
        
            vecteur(1) = x
        
            vecteur(2) = y

        ResoudreGenerale = vecteur
' Aucun résultat possible pour un nombre d'itérations inferieur à 1
    Else
        ResoudreGenerale = CVErr(xlErrNum)
    
    End If
    
End Function


