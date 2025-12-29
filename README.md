
# VBA Iterative Linear System Solver / Résolveur Itératif de Systèmes Linéaires VBA

## Project Information / Informations sur le Projet

**Course**: P.M.Q.F. 20671 - HEC Montréal  
**Team**: Alex-Anne Favreau (11297786), Julia Krumgant (11298595)

This project demonstrates VBA programming skills including function creation, iterative algorithms, array handling, and Excel integration.

---

## VBA Skills Demonstrated / Compétences VBA Démontrées

### 1. Custom Functions / Fonctions Personnalisées
- Created user-defined functions callable from Excel worksheets
- Implemented mathematical formulas in VBA syntax
- Proper function declaration with parameters and return types

### 2. Algorithm Implementation / Implémentation d'Algorithme
- Fixed-point iteration method (Gauss-Seidel style)
- Loop control structures (`For...Next`)
- Conditional logic for error handling

### 3. Array Operations / Opérations sur les Tableaux
- Return arrays from functions to Excel
- Array indexing and manipulation
- Multi-value returns

### 4. Excel Integration / Intégration Excel
- Functions callable directly in cells: `=ResoudreGenerale(...)`
- Accept Range parameters from worksheets
- Return values displayed in Excel cells

---

## Code Structure / Structure du Code

**Three Main Functions:**

1. **`Eq1IsoleXGenerale(y, a1, b1, c1)`** - Isolates x from the first equation
2. **`Eq1IsoleYGenerale(x, a2, b2, c2)`** - Isolates y from the second equation  
3. **`ResoudreGenerale(y0, n, a1, b1, c1, a2, b2, c2)`** - Main solver with iteration loop

**Key VBA Code Pattern:**
```vba
For i = 1 To n
    x = Eq1IsoleXGenerale(y, a1, b1, c1)
    y = Eq1IsoleYGenerale(x, a2, b2, c2)
Next i

Dim vecteur(1 To 2)
vecteur(1) = x
vecteur(2) = y
ResoudreGenerale = vecteur
```

---

## Usage Example / Exemple d'Utilisation

**In Excel cell:**
```excel
=ResoudreGenerale(0, 10, 2, 3, 13, 4, 1, 11)
```

Returns the solution as an array after 10 iterations.

---

## Key VBA Concepts Applied / Concepts VBA Appliqués

✓ Function procedures with `Function...End Function`  
✓ Variable declaration with `Dim`  
✓ Parameter passing  
✓ Iterative loops with `For...Next`  
✓ Conditional statements with `If...Then...Else`  
✓ Array creation and return  
✓ Error handling with `CVErr(xlErrNum)`  
✓ Mathematical operations  
✓ Excel Range integration

---

## Technical Specifications

- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Microsoft Excel
- **Method**: Iterative solver for 2x2 linear systems
- **Input**: System coefficients and iteration parameters
- **Output**: Solution array [x, y]
