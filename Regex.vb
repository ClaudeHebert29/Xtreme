Imports System.Text.RegularExpressions

Public Class clsRegex

    Dim regexTelephone As New Regex("^(\((\d{3})\)|(\d{3}))\s*[-\/\.]?\s*(\d{3})\s*[-\/\.]?\s*(\d{4})$")
    Dim regexTelFrance As New Regex("^(0[1-68])([ _.\-]?(\d{2})){4}$")
    Dim regexPosteTel As New Regex("^\d{1,5}$")
    Dim regexCouriel As New Regex("^[\.a-z0-9_\-]+@[a-z0-9_-]+(\.[a-z]{2,4}){1,2}$")
    Dim regexNomPrenom As New Regex("^[-a-zA-Z 'ÇüéâäàåçêëèïîìÁÂÀÄÅÉÈÊæÆôöòûùÿÔÒÓÖÙÛÚÜáíóúñÑ]{1,50}$")
    Dim regexLetterOnly As New Regex("^[a-zA-ZéÉèÈêÊ \-']+$")
    Dim regexCodePostal As New Regex("^[ABCEGHJKLMNPRSTVXY]{1}\d{1}[ABCEGHJKLMNPRSTVWXYZ]{1}(\ |)\d{1}[ABCEGHJKLMNPRSTVWXYZ]{1}\d{1}$")
    Dim regexCodePostalFrance As New Regex("^[\d]{5}$")
    Dim regexAdresse As New Regex("^[\,\(\)0-9a-zA-ZéÉèÈêÊ \-]{1,50}$")
    Dim regexAssMal As New Regex("^[a-zA-Z]{4}[0-9]{2}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])[0-9]{2}$")
    Dim regexNumberOnly As New Regex("^[\d]{1,50}$")
    Dim regexNumber As New Regex("^[\d]{1,3}([\.,][\d]{2})?$")
    Dim regexExpiration As New Regex("^((1[012])|(0[1-9]))\/[\d]{2}$")
    Dim regexHeure As New Regex("^(([0-1][0-9])|([2][0-3]))h[0-5][0-9]$")
    Dim regexInitiale As New Regex("^[A-Za-z]{2,5}$")
    Dim regexBraquet As New Regex("[\{\}\[\]<>]")

    Public Function checkNom(ByVal _nom As String)
        Return regexNomPrenom.IsMatch(_nom)
    End Function
    Public Function checkTel(ByVal _tel As String)
        Dim _valide As Boolean = False
        _valide = regexTelephone.IsMatch(_tel)
        If _valide = False Then
            _valide = regexTelFrance.IsMatch(_tel)
        End If
        Return _valide
    End Function
    Public Function checkCouriel(ByVal _couriel As String)
        Return regexCouriel.IsMatch(_couriel)
    End Function
    Public Function checkPosteTel(ByVal _poste As String)
        Return regexPosteTel.IsMatch(_poste)
    End Function
    Public Function checkLetterOnly(ByVal _text As String)
        Dim _valide As Boolean = False
        _valide = regexLetterOnly.IsMatch(_text)
        Return _valide
    End Function
    Public Function checkCodePostal(ByVal _codePostal As String)
        Dim _valide As Boolean = False
        _valide = regexCodePostal.IsMatch(_codePostal)
        If _valide = False Then
            _valide = regexCodePostalFrance.IsMatch(_codePostal)
        End If
        Return _valide
    End Function
    Public Function checkAdresse(ByVal _adresse As String)
        Return regexAdresse.IsMatch(_adresse)
    End Function
    Public Function checkAssMal(ByVal _assMal As String)
        Return regexAssMal.IsMatch(_assMal)
    End Function
    Public Function checkNumber(ByVal _number As String)
        Return regexNumber.IsMatch(_number)
    End Function
    Public Function checkExpiration(ByVal _Expiration As String)
        Return regexExpiration.IsMatch(_Expiration)
    End Function
    Public Function checkHeure(ByVal _heure As String)
        Return regexHeure.IsMatch(_heure)
    End Function
    Public Function checkInitiale(ByVal _ini As String)
        Return regexInitiale.IsMatch(_ini)
    End Function

    Public Function checkBraquet(ByVal _texte As String)
        Return Not regexBraquet.IsMatch(_texte)
    End Function
End Class
