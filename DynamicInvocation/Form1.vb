Imports System
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DI As New DynamicInvoke
        DI.Invoke("MessageBox", "User32", GetType(Int16), 0, "Ola", "App", 0)

        SendKeys.Send()
        End
    End Sub
End Class



Public Class DynamicInvoke
    ',  Optional ByVal AssemblyName As String = "DynamicInvoke", Optional ByVal TypeName As String = "DynamicType", Optional ByVal Convention As CallingConvention = CallingConvention.Winapi, Optional ByVal CharacterSet As CharSet = CharSet.Ansi

    Public AssemblyName As String = "DynamicInvoke"
    Public TypeName As String = "DynamicType"
    Public Convention As CallingConvention = CallingConvention.Winapi
    Public CharacterSet As CharSet = CharSet.Ansi

    Function Invoke(ByVal MethodName As String, ByVal LibraryName As String, ByVal ReturnType As Type, ByVal ParamArray Parameters() As Object)

        Dim ParameterTypesArray As Array

        If Parameters IsNot Nothing Then
            ParameterTypesArray = Array.CreateInstance(GetType(Type), Parameters.Length)
            Dim PTIndex As Integer = 0

            For Each Item In Parameters
                If Item IsNot Nothing Then
                    ParameterTypesArray(PTIndex) = Item.GetType
                Else
                    'ParameterTypesArray(PTIndex) = 0
                End If
                PTIndex += 1
            Next
        Else
            ParameterTypesArray = Nothing
        End If

        Dim ParameterTypes() As Type = ParameterTypesArray


        Dim asmName As New AssemblyName(AssemblyName)
        Dim dynamicAsm As AssemblyBuilder = _
            AppDomain.CurrentDomain.DefineDynamicAssembly(asmName, _
                AssemblyBuilderAccess.RunAndSave)

        ' Create the module.
        Dim dynamicMod As ModuleBuilder = _
            dynamicAsm.DefineDynamicModule(asmName.Name, asmName.Name & ".dll")

        ' Create the TypeBuilder for the class that will contain the 
        ' signature for the PInvoke call.
        Dim tb As TypeBuilder = dynamicMod.DefineType(TypeName, _
            TypeAttributes.Public Or TypeAttributes.UnicodeClass)

        Dim mb As MethodBuilder = tb.DefinePInvokeMethod( _
            MethodName, _
            LibraryName, _
            MethodAttributes.Public Or MethodAttributes.Static Or MethodAttributes.PinvokeImpl, _
            CallingConventions.Standard, _
            ReturnType, _
            ParameterTypes, _
            Convention, _
            CharacterSet)

        ' Add PreserveSig to the method implementation flags. NOTE: If this line
        ' is commented out, the return value will be zero when the method is
        ' invoked.
        mb.SetImplementationFlags( _
            mb.GetMethodImplementationFlags() Or MethodImplAttributes.PreserveSig)

        ' The PInvoke method does not have a method body.

        ' Create the class and test the method.
        Dim t As Type = tb.CreateType()

        Dim mi As MethodInfo = t.GetMethod(MethodName)
        Return mi.Invoke(Me, Parameters)

        '' Produce the .dll file.
        'Console.WriteLine("Saving: " & asmName.Name & ".dll")
        'dynamicAsm.Save(asmName.Name & ".dll")
    End Function

End Class
