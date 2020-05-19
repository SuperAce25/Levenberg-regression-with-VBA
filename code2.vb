'Bibliography: A brief description of the Levenberg-Marquadt Algorithm Implemented by levmar, Lourakis M., 2005, Institute of Computer Science
'This sub apply the levenberg-marquadt algorithm for nonlinear least squares regression, also could apply Weigthed least squares
'This regression is adapted to three parameters but can also be adapted to a different number of parameters
Function Levmar(X, y, F, dfu, initial_theta, Optional ByVal chain_param As Integer = -999, Optional ByVal weighted As Boolean = False, Optional accel As Boolean = False) As Double()
'X is a vector with m examples, X can be a multidimensional vector but the problem has to be modified
'y is the result vector of the same dimension as X
'F is the function to be regressed, this is called as an User Defined Function, so you have to create a Function program and then put its name in F
'dfu is the conjunction of derivatives of F, if F has n parameters then it has to have n condition for each parameter
'dfu is called as an UDF with the partial derivative indicated in the last parameter of the UDF
'i.e. F=a*x+b*x^2+c*x^3, then dfu(a,b,c,X,1)=x; dfu(a,b,c,X,2)=2*b*x; dfu(a,b,c,X,3)=3*c*x^2, where 1,2 and 3 inside dfu indicate the partial derivative of the parameters a,b,c
'If you didn't understand I invite you to see how the function work
'initial_theta are the initial parameters of the function F, for example [1 2 3] can be the initial parameters of (a,b,c) of the function F=a*x+b*x^2+c*x^3
'weighted is an option if you want to calculate the Weighted Non-linear least squares where the weights add more importance to some values of X and y than others, this is used if you don't have reliable data
'The weights are calculated from the number of points that are in an interval in X, if X is a class-interval data
'If you want to change the weights of each data, or if you now the weights then you have to create a matrix mxm and add it to w()
Application.ScreenUpdating = False 'We shut down the screen updating
Dim m As Integer, n As Integer
m = UBound(X) - LBound(X) + 1 'Number of points
n = UBound(initial_theta) 'number of parameters
n_min = LBound(initial_theta)
max_iters = 200 'Number of maximum iterations
Dim lambda As Double 'Regularization term
Dim Weight As Boolean
ReDim N_Interval(1 To m, 1 To 1) As Double 'Number of points of every interval in X
ReDim theta(LBound(initial_theta) To UBound(initial_theta)) As Double 'parameters vector
ReDim theta2(LBound(initial_theta) To UBound(initial_theta)) As Double 'parameters vector for geodesic acc
ReDim J_history(0 To max_iters) As Double 'Cost function vector
ReDim J(1 To m, 1 To n) As Double 'Jacobian
ReDim r(1 To m, 1 To 1) As Double 'residual vector
ReDim r2(1 To m, 1 To 1) As Double 'residual vector for geodesic acc
ReDim a3(1 To m, 1 To 1) As Double
w = (ones(m, 1)) 'covariance matrix mxm
eps1 = 0.0000000001 'Error of Jacobian and objective function JT*dB (See wikipedia)
eps2 = 0.0000000001 'Error of parameter vector dtheta
eps3 = 0.0000000001 'Total error dBt*dB (dB is the difference between the example data and the objective function)
tau = 0.001 'For inicialization of lambda
V = 2 'Step size increment for lamba
h = 0.1 'finite difference step size for the second directional derivative for the geodesic acceleration

'inicialization of the parameters vector
theta = initial_theta

'Iteration 0
For i = 1 To m
    r(i, 1) = y(i) - Application.Run(F, theta, X(i)) 'Here we calculate the error of y-F(X)
    For k = 1 To n
        J(i, k) = Application.Run(dfu, theta, X(i), k) 'k is the derivated parameter <- Here is the Jacobian
    Next k
    If weighted = True Then
        If r(i, 1) = 0 Then
            w(i, i) = 1
        Else
            w(i, i) = (m - n + 1) / (r(i, 1) ^ 2)
        End If
    End If
Next i
'w(i, i) = 1 / ((y(i) + Abs(r(i, 1))) * N_c(i)) where Nc are the numbers of points in the bin i

'Transposition of the Jacobian
JT = WorksheetFunction.Transpose(J)
'if the Levmar is weighted then multiply the Jt by the weight matrix
If weighted = True Then
    JT = WorksheetFunction.MMult(JT, w)
End If

DF2 = WorksheetFunction.MMult(JT, J) 'nxn matrix <- Aproximation of second derivatives of F

'Initialization of lambda
For k = 1 To n
If DF2(k, k) > lambda Then
lambda = DF2(k, k)
End If
Next k

lambda = tau * lambda

B = WorksheetFunction.MMult(JT, r) 'nx1 matrix
'First and third condition
e1 = inf_norm(B) '<- we find the inf-norm of B (I think)
e3 = magnitude(r) 'another type of error
'Calculation of the first cost function
J_history(0) = e3 ^ 2
min_cost = J_history(0)

If e1 > eps1 And e3 ^ 2 > eps3 Then 'if the maximum value of Jt*r is small or the residual magnitude is small then stop

    For iters = 1 To max_iters
    
        For k = 1 To n
            If Abs(DF2(k, k)) < 1 Then
                DF2(k, k) = DF2(k, k) + lambda
            Else
                DF2(k, k) = DF2(k, k) + lambda * DF2(k, k) '<- We add the regularization term (This is the part of the algorithm that make it different from the Gauss-Newton)
            End If
        Next k
        
        dtheta = WorksheetFunction.MMult(WorksheetFunction.MInverse(DF2), B) 'mx1 matrix, we solve the system ((Jt*W*J+lambda*diag(Jt*W*J))^-1)*Jt*W*r
        
        'Geodesic acceleration('Unstable')
        If accel = True Then
            For k = n_min To n
                If chain_param <> 999 And k = chain_param Then GoTo Nextiteration1
                theta2(k) = theta(k) + h * dtheta(k, 1) '<-We replace the old parameters for the new parameters if the errors of the iteration is still greater than the error defined
Nextiteration1:
            Next k
            J_theta = WorksheetFunction.MMult(J, dtheta)
            For i = 1 To m
                r2(i, 1) = (2 / h) * ((((y(i) - Application.Run(F, theta2, X(i))) - r(i, 1)) / h) - J_theta(i, 1))
            Next i
    
            b2 = WorksheetFunction.MMult(JT, r2)
            'b2 = WorksheetFunction.MMult(WorksheetFunction.Transpose(J), r2)
            dtheta2 = WorksheetFunction.MMult(WorksheetFunction.MInverse(DF2), b2)
            
            alfa = magnitude(dtheta2) / magnitude(dtheta)
            
            If alfa <= 0.75 Then
                For k = 1 To n
                    dtheta(k, 1) = dtheta(k, 1) + 0.5 * dtheta2(k, 1)
                Next k
            End If
        
        End If
        'End of Geodesic acceleration
        
        'Second condition
        e2 = magnitude(dtheta)
        e21 = magnitude(theta)
        
        If e2 <= eps2 * e21 Then 'if the change in dtheta is small compared to theta then stop
            Exit For
        Else
            For k = 1 To n
                theta(k) = theta(k) + dtheta(k, 1)
            Next k
        End If
        
        'We compute the general error of the objective function with parameters theta
        
        For i = 1 To m
            r2(i, 1) = y(i) - Application.Run(F, theta, X(i))
        Next i
        
        J_history(iters) = magnitude(r2) ^ 2
        
        For k = 1 To n
            rho = rho + dtheta(k, 1) * (lambda * DF2(k, k) * dtheta(k, 1) + B(k, 1))
        Next k
    
        rho = (min_cost - J_history(iters)) / rho
        
        'We assure that the new theta is better than before
        If rho > 0 Then
            lambda = lambda * WorksheetFunction.Max(1 / 3, (1 - (2 * rho - 1) ^ 3))
            V = 2
            min_cost = J_history(iters)
            For i = 1 To m
                r(i, 1) = y(i) - Application.Run(F, theta, X(i)) 'Here we calculate the error of y-F(X)
                For k = n_min To n
                    J(i, k) = Application.Run(dfu, theta, X(i), k) 'k is the derivated parameter <- Here is the Jacobian
                Next
                If weighted = True Then
                    If r(i, 1) = 0 Then
                        w(i, i) = 1
                    Else
                        w(i, i) = (m - n + 1) / (r(i, 1) ^ 2)
                    End If
                End If
            Next
            
            JT = WorksheetFunction.Transpose(J)
        
            If weighted = True Then
                JT = WorksheetFunction.MMult(JT, w)
            End If
            
            B = WorksheetFunction.MMult(JT, r) 'nx1 matrix
            
            'First and third condition
            e1 = inf_norm(B) '<- we find the inf-norm of B (I think)
            e3 = magnitude(r) 'another type of error
            
            If e1 <= eps1 Or e3 ^ 2 <= eps3 Then 'if the maximum value of Jt*r is small or the residual magnitude is small then stop
                Exit For
            End If
            
            DF2 = WorksheetFunction.MMult(JT, J) 'nxn matrix <- Aproximation of second derivatives of F
            
        Else
            lambda = V * lambda
            V = 2 * V
            For k = n_min To n
                If chain_param <> 999 And k = chain_param Then GoTo pasa
                theta(k) = theta(k) - dtheta(k, 1)
pasa:
            Next
            
        End If
        
    Next iters
End If
'Copy the parameters to excel
Range("M2").Value = e3 ^ 2 / m 'Mean squared error
Range("M3").Value = iters
Application.ScreenUpdating = True

Levmar = theta

End Function
