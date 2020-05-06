# Levenberg-regression-with-VBA
Nonlinear weighted regression with Levenberg-Marquadt algorithm in VBA
'This sub apply the levenberg-marquadt algorithm for nonlinear least squares regression, also could apply Weigthed least squares
'This regression is adapted to three parameters but can also be adapted to a different number of parameters
Sub Levmar(X, y, F, dfu, Jcost, initial_theta, weighted)
'X is a vector with m examples, X can be a multidimensional vector but the problem has to be modified
'y is the result vector of the same dimension as X
'F is the function to be regressed, this is called as an User Defined Function, so you have to create a Function program and then put its name in F
'dfu is the conjunction of derivatives of F, if F has n parameters then it has to have n condition for each parameter
'dfu is called as an UDF with the partial derivative indicated in the last parameter of the UDF
'i.e. F=a*x+b*x^2+c*x^3, then dfu(a,b,c,X,1)=x; dfu(a,b,c,X,2)=2*b*x; dfu(a,b,c,X,3)=3*c*x^2, where 1,2 and 3 inside dfu indicate the partial derivative of the parameters a,b,c
'If you didn't understand I invite you to see how the function work
'Jcost is a UDF that calculates the error from the calculation of F(X) and y
'initial_theta are the initial parameters of the function F, for example [1 2 3] can be the initial parameters of (a,b,c) of the function F=a*x+b*x^2+c*x^3
'weighted is an option if you want to calculate the Weighted Non-linear least squares where the weights add more importance to some values of X and y than others, this is used if you don't have reliable data
'The weights are calculated from the number of points that are in an interval in X, if X is a class-interval data
'If you want to change the weights of each data, or if you now the weights then you have to create a matrix mxm and add it to w()

Application.ScreenUpdating = False
Dim m As Integer, n As Integer
m = UBound(X) - LBound(X) + 1 'Number of points
n = UBound(initial_theta) - LBound(initial_theta) + 1 'Number of parameters
num_iters = 100 'Number of maximum iterations
Dim lambda As Double 'Regularization term
Dim Weight As Boolean
ReDim N_Interval(1 To m, 1 To 1) As Double 'Number of points of every interval in X
ReDim theta(0 To n - 1) As Double
ReDim dJ(1 To n) As Double
ReDim J_history(1 To num_iters) As Double
ReDim df(1 To m, 1 To n) As Double
ReDim dC(1 To m, 1 To 1) As Double
Weight = weighted
w = (ones(m, 1))
eps1 = 0.0000000001 'Error of Jacobian and objective function Jt*dB (See wikipedia)
eps2 = 0.0000000001 'Error of parameter vector dtheta
eps3 = 0.0000000001 'Total error dBt*dB (dB is the difference between the example data and the objective function)
tau = 0.001 'For inicialization of lambda (Not usable for VBA because of non-array functions)
lambda = 1
V = 2


'This is the initial aproximation (As always if this is far away from the optimum then the algorithm will diverge)
'Usually happens with the Power Function (is not a good algorithm for that function in particular)
Nugget = ActiveWorkbook.Sheets("Variogram Model").Range("J3")
'For this algorithm is just for 3 parameters, but you can modify to add more parameters, watch out for other changes!
theta(0) = initial_theta(1)
theta(1) = initial_theta(2)
theta(2) = initial_theta(3)

If Nugget = "True" Then '<-This part is important, you can chain a parameter and calculate the rest
    n_min = 1
Else
    theta(0) = ActiveWorkbook.Sheets("Variogram Model").Range("J4") 'This is if you want to chained one of the parameters
    n_min = 2 'You can chained various parameters but you have to modify the code
End If

For i = 1 To m
    N_Interval(i, 1) = ActiveWorkbook.Sheets("Data").Cells(24 + i, 10) 'This is for the weighted least squares method, it can be turn off
Next

For iters = 1 To num_iters
    For i = 1 To m
        dC(i, 1) = y(i) - Application.Run(F, theta, X(i)) 'Here we calculate the error of y-F(X)
        For k = n_min To n
            df(i, k) = Application.Run(dfu, theta, X(i), k - 1) 'k-1 is the derivated parameter <- Here is the Jacobian
        Next
        If Weight = True Then
            If X(i) = 0 Then
                w(i, i) = 1
            Else
                w(i, i) = N_Interval(i, 1) / (Application.Run(F, theta, X(i)) ^ 2) '<-Here you add the weighted matrix
            End If
        End If
    Next
    df_transpose = WorksheetFunction.Transpose(df)
    
    If Weight = True Then
        df_transpose = WorksheetFunction.MMult(df_transpose, w)
    End If
    
    DF2 = WorksheetFunction.MMult(df_transpose, df) 'nxn matrix <- Aproximation of second derivatives of F
    For k = 1 To n
        DF2(k, k) = DF2(k, k) + lambda '<- We add the regularization term (This is the part of the algorithm that make it different from the Gauss-Newton)
    Next
    B = WorksheetFunction.MMult(df_transpose, dC) 'nx1 matrix
    'First and third condition
    e1 = 0 '<- we find the 1-norm of B (I think)
    For Each Item In B
        If Abs(Item) > e1 Then
        e1 = Abs(Item)
        End If
    Next
    e3 = magnitude(B) 'another type of error
    If e1 <= eps1 Or e3 ^ 2 <= eps3 Then
        Exit For
    End If
    dtheta = WorksheetFunction.MMult(WorksheetFunction.MInverse(DF2), B) 'mx1 matrix '<- If excel wouldn't have this, we would have to create a method for linear equations systems, like Jacobi, Gauss-Siedel or others.
    'Second condition
    e2 = magnitude(dtheta)
    e21 = magnitude(theta)
    If e2 <= eps2 * e21 Then '<- another one
        Exit For
    End If
    
    For k = n_min To n
        theta(k - 1) = theta(k - 1) + dtheta(k, 1) '<-We replace the old parameters for the new parameters if the errors of the iteration is still greater than the error defined
    Next
    'We compute the general error of the objective function with parameters theta
    J_history(iters) = Application.Run(Jcost, X, y, theta, m)
    'Here comes the damped iteration, this make the iteration run more slower or faster
    If iters > 1 Then
        If J_history(iters) < J_history(iters - 1) Then
            lambda = lambda / V
        Else
            lambda = V * lambda
            For k = n_min To n
                theta(k - 1) = theta(k - 1) - dtheta(k, 1)
            Next
        End If
    End If
Next

For kl = 1 To 3
    Cells(kl + 18, 8) = theta(kl - 1)
Next

Range("M2") = J_history(iters - 1)
Range("M3") = iters - 1
MsgBox "Fin"
Application.ScreenUpdating = True
End Sub
