# VBA Linear System Solver / Résolveur de Systèmes Linéaires VBA

## English Version

### Project Overview

This project implements a comprehensive linear system solver in VBA (Visual Basic for Applications) for Excel. It provides tools to solve systems of linear equations using various numerical methods, with built-in validation, error handling, and user-friendly interfaces.

### Features

- **Multiple Solution Methods**: Implements Gaussian elimination with partial pivoting
- **Automatic System Validation**: Checks for system compatibility and consistency
- **Error Management**: Robust error handling for division by zero, incompatible systems, and numerical precision issues
- **Excel Integration**: Seamlessly works with Excel ranges and outputs results directly to worksheets
- **User-Defined Functions**: Custom Excel functions for easy integration into spreadsheets
- **Interactive Interface**: User-friendly forms for inputting systems and viewing solutions

### Files Description

#### Macros VBA Resolver.bas

This is the main VBA module containing all the core functionality:

**Main Functions:**

1. **`SolveLinearSystem(A As Range, b As Range) As Variant`**
   - Solves a linear system Ax = b
   - Parameters:
     - `A`: Matrix of coefficients (m x n range)
     - `b`: Vector of constants (m x 1 range)
   - Returns: Solution vector x or error message
   - Usage: Can be called as a worksheet function `=SolveLinearSystem(A1:C3, D1:D3)`

2. **`GaussianElimination(matrix As Variant) As Variant`**
   - Performs Gaussian elimination with partial pivoting
   - Transforms the augmented matrix to row echelon form
   - Handles pivot selection to minimize numerical errors

3. **`BackSubstitution(matrix As Variant) As Variant`**
   - Performs back substitution on an upper triangular matrix
   - Extracts the solution vector from the reduced matrix

4. **`ValidateSystem(A As Variant, b As Variant) As Boolean`**
   - Validates input matrices for compatibility
   - Checks dimensions and data types
   - Returns True if system is valid

5. **`CheckConvergence(matrix As Variant) As String`**
   - Analyzes the system for:
     - Unique solutions (non-singular systems)
     - Infinite solutions (underdetermined systems)
     - No solutions (inconsistent systems)
   - Returns: "UNIQUE", "INFINITE", or "NONE"

**Helper Functions:**

- **`SwapRows(matrix As Variant, row1 As Integer, row2 As Integer)`**: Exchanges two rows in the matrix
- **`FindPivot(matrix As Variant, col As Integer) As Integer`**: Locates the best pivot element in a column
- **`IsZeroRow(matrix As Variant, row As Integer) As Boolean`**: Checks if a row contains all zeros
- **`FormatSolution(solution As Variant) As String`**: Formats the solution vector for display

### Installation and Usage

#### Installation Steps:

1. **Open Excel** and press `Alt + F11` to access the VBA Editor
2. **Import the Module**: Go to `File > Import File` and select `Macros VBA Resolver.bas`
3. **Enable Macros**: Ensure macros are enabled in your Excel security settings
4. **Save as Macro-Enabled Workbook**: Save your file with the `.xlsm` extension

#### Usage Examples:

**Example 1: Using as a Worksheet Function**

```excel
' Setup your system:
' A1:C3 contains the coefficient matrix:
'   2   1  -1
'  -3  -1   2
'  -2   1   2
'
' D1:D3 contains the constants vector:
'   8
'  -11
'  -3
'
' In cell F1, enter:
=SolveLinearSystem(A1:C3, D1:D3)
```

**Example 2: Using VBA Code**

```vba
Sub SolveMySystem()
    Dim A As Range
    Dim b As Range
    Dim solution As Variant
    
    ' Define your system
    Set A = Range("A1:C3")
    Set b = Range("D1:D3")
    
    ' Solve the system
    solution = SolveLinearSystem(A, b)
    
    ' Output results
    Range("F1").Resize(UBound(solution), 1).Value = solution
End Sub
```

**Example 3: Error Handling**

The solver automatically detects and reports:
- Division by zero situations
- Incompatible matrix dimensions
- Singular matrices (no unique solution)
- Numerical precision issues

### Technical Specifications

#### Algorithm Details:

- **Method**: Gaussian Elimination with Partial Pivoting
- **Complexity**: O(n³) for an n×n system
- **Precision**: Double precision floating-point arithmetic
- **Stability**: Partial pivoting ensures numerical stability

#### System Requirements:

- Microsoft Excel 2010 or later
- VBA 7.0 or later
- Macros must be enabled

#### Limitations and Considerations:

- **Convergence**: Not verified explicitly for all cases - assumes well-conditioned systems
- **Division by Zero**: Handled through pivot selection, but near-zero pivots may cause instability
- **System Compatibility**: Well-adapted for most standard linear systems
- **Numerical Precision**: Results are accurate for most practical applications but may have limitations with very large or ill-conditioned systems

### Educational Context

This project was created in an educational context for:
- Translating algebraic equations into code
- Implementing iterative numerical methods
- Creating user-defined functions in Excel
- Managing arrays and tables in VBA
- Returning results from VBA to Excel
- Implementing simple error handling

### Project Repository

**Owner / Propriétaire**: Julia11614

---

## Version Française

### Aperçu du Projet

Ce projet implémente un résolveur complet de systèmes linéaires en VBA (Visual Basic for Applications) pour Excel. Il fournit des outils pour résoudre des systèmes d'équations linéaires en utilisant diverses méthodes numériques, avec validation intégrée, gestion des erreurs et interfaces conviviales.

### Fonctionnalités

- **Plusieurs Méthodes de Résolution**: Implémente l'élimination de Gauss avec pivotage partiel
- **Validation Automatique du Système**: Vérifie la compatibilité et la cohérence du système
- **Gestion des Erreurs**: Gestion robuste des erreurs pour division par zéro, systèmes incompatibles et problèmes de précision numérique
- **Intégration Excel**: Fonctionne parfaitement avec les plages Excel et affiche les résultats directement dans les feuilles de calcul
- **Fonctions Définies par l'Utilisateur**: Fonctions Excel personnalisées pour une intégration facile dans les feuilles de calcul
- **Interface Interactive**: Formulaires conviviaux pour saisir les systèmes et visualiser les solutions

### Description des Fichiers

#### Macros VBA Resolver.bas

Ceci est le module VBA principal contenant toutes les fonctionnalités de base:

**Fonctions Principales:**

1. **`SolveLinearSystem(A As Range, b As Range) As Variant`**
   - Résout un système linéaire Ax = b
   - Paramètres:
     - `A`: Matrice des coefficients (plage m x n)
     - `b`: Vecteur des constantes (plage m x 1)
   - Retourne: Vecteur solution x ou message d'erreur
   - Utilisation: Peut être appelée comme fonction de feuille de calcul `=SolveLinearSystem(A1:C3, D1:D3)`

2. **`GaussianElimination(matrix As Variant) As Variant`**
   - Effectue l'élimination de Gauss avec pivotage partiel
   - Transforme la matrice augmentée en forme échelonnée réduite
   - Gère la sélection du pivot pour minimiser les erreurs numériques

3. **`BackSubstitution(matrix As Variant) As Variant`**
   - Effectue la substitution arrière sur une matrice triangulaire supérieure
   - Extrait le vecteur solution de la matrice réduite

4. **`ValidateSystem(A As Variant, b As Variant) As Boolean`**
   - Valide les matrices d'entrée pour la compatibilité
   - Vérifie les dimensions et les types de données
   - Retourne True si le système est valide

5. **`CheckConvergence(matrix As Variant) As String`**
   - Analyse le système pour:
     - Solutions uniques (systèmes non singuliers)
     - Solutions infinies (systèmes sous-déterminés)
     - Aucune solution (systèmes incohérents)
   - Retourne: "UNIQUE", "INFINITE", ou "NONE"

**Fonctions Auxiliaires:**

- **`SwapRows(matrix As Variant, row1 As Integer, row2 As Integer)`**: Échange deux lignes dans la matrice
- **`FindPivot(matrix As Variant, col As Integer) As Integer`**: Localise le meilleur élément pivot dans une colonne
- **`IsZeroRow(matrix As Variant, row As Integer) As Boolean`**: Vérifie si une ligne ne contient que des zéros
- **`FormatSolution(solution As Variant) As String`**: Formate le vecteur solution pour l'affichage

### Installation et Utilisation

#### Étapes d'Installation:

1. **Ouvrir Excel** et appuyer sur `Alt + F11` pour accéder à l'éditeur VBA
2. **Importer le Module**: Aller à `Fichier > Importer un fichier` et sélectionner `Macros VBA Resolver.bas`
3. **Activer les Macros**: S'assurer que les macros sont activées dans les paramètres de sécurité Excel
4. **Enregistrer comme Classeur avec Macros**: Enregistrer votre fichier avec l'extension `.xlsm`

#### Exemples d'Utilisation:

**Exemple 1: Utilisation comme Fonction de Feuille de Calcul**

```excel
' Configurer votre système:
' A1:C3 contient la matrice des coefficients:
'   2   1  -1
'  -3  -1   2
'  -2   1   2
'
' D1:D3 contient le vecteur des constantes:
'   8
'  -11
'  -3
'
' Dans la cellule F1, entrer:
=SolveLinearSystem(A1:C3, D1:D3)
```

**Exemple 2: Utilisation du Code VBA**

```vba
Sub ResoudreMonSysteme()
    Dim A As Range
    Dim b As Range
    Dim solution As Variant
    
    ' Définir votre système
    Set A = Range("A1:C3")
    Set b = Range("D1:D3")
    
    ' Résoudre le système
    solution = SolveLinearSystem(A, b)
    
    ' Afficher les résultats
    Range("F1").Resize(UBound(solution), 1).Value = solution
End Sub
```

**Exemple 3: Gestion des Erreurs**

Le résolveur détecte et signale automatiquement:
- Situations de division par zéro
- Dimensions de matrice incompatibles
- Matrices singulières (pas de solution unique)
- Problèmes de précision numérique

### Spécifications Techniques

#### Détails de l'Algorithme:

- **Méthode**: Élimination de Gauss avec Pivotage Partiel
- **Complexité**: O(n³) pour un système n×n
- **Précision**: Arithmétique en virgule flottante double précision
- **Stabilité**: Le pivotage partiel assure la stabilité numérique

#### Exigences Système:

- Microsoft Excel 2010 ou ultérieur
- VBA 7.0 ou ultérieur
- Les macros doivent être activées

#### Limites et Considérations:

- **Convergence**: Non vérifiée explicitement pour tous les cas - suppose des systèmes bien conditionnés
- **Division par Zéro**: Gérée par la sélection du pivot, mais des pivots proches de zéro peuvent causer de l'instabilité
- **Compatibilité Système**: Bien adapté pour la plupart des systèmes linéaires standards
- **Précision Numérique**: Les résultats sont précis pour la plupart des applications pratiques mais peuvent avoir des limitations avec des systèmes très grands ou mal conditionnés

### Contexte Pédagogique

Ce projet a été réalisé dans le cadre d'un contexte éducatif pour:
- Traduction d'équations algébriques en code
- Implémentation de méthodes numériques itératives
- Création de fonctions définies par l'utilisateur dans Excel
- Gestion de tableaux et de tables en VBA
- Retour de tableaux depuis VBA vers Excel
- Gestion simple des erreurs

### Propriétaire du Dépôt

**Repository Owner**: Julia11614
