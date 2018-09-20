Imports System.Net
Imports System.Net.Mail
'EmailNotify.SendEmail("headletters10 Logic Test for (ALL) " + ST.ToString(), ST, "end")
Module EmailNotify
    Public ConsoleLogs = ""
    Sub SendEmail(ByVal Subject As String, ByVal Time As DateTime, ByVal flag As String)
        Dim mail As New MailMessage
        mail.To.Add("abhishek.singh@enukesoftware.com")
        'mail.To.Add("vcdubai@gmail.com")
        mail.From = New MailAddress("iamemailsender@gmail.com")
        mail.Subject = Subject
        mail.IsBodyHtml = True
        If flag = "start" Then
        End If
        If flag = "end" Then
            mail.Body += $"<!DOCTYPE html>
<html lang='en'>

<head>
    <meta charset='UTF-8'>
    <title>Chart</title>
    <link href='https//fonts.googleapis.com/css?family=Roboto+Slab' rel='stylesheet'>   
</head>
<style>
</style>
<body>

    <div style='width850px;margin:0 auto;border:1px soid #e3e3e3;padding:15px;background: #f5f5f8;border-radius: 3px;box-shadow: 0 3px 5px 0px #E1DAD1;'>

    <table style='margin: 10px auto 20px;background:#fff;padding: 40px 30px;'>
        <tr>
            <td>
                <table width='744'>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Name of the Chart Owner :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>Shubham</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Date of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>20-08-1993</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Place of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>New Delhi</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Time of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>9:05 AM</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Latitude :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>34534553</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Longtitude :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>67657567</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Day of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>Sunday</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Rashi :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>WHat</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Star :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>Monday</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Sun Sign :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>7:20 AM</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Balance Dasa :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>34534</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style='margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
            <td>
                <h3 class='box-head' style='margin: 0 0 20px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>Planetary position</h3></td>
        </tr>
        <tr>
            <td>
                <table width='744' style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>
                    <tr>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Planet</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sign</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D-M-S</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>RL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>STL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SSL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SukL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>PraL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Hse</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>PS</th>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>I</td>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>I</td>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>I</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style='margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
            <td>
                <h3 class='box-head' style='margin: 0 0 20px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>Cuspal position</h3></td>
        </tr>
        <tr>
            <td>
                <table width='744' style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>
                    <tr>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Cusp</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sign</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D-M-S</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>RL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>STL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SSL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>SukL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>PraL</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>%SArc</th>
                        <th style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>%PArc</th>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                    </tr>
                    <tr>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>B</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>C</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>D</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>E</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>F</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>G</td>
                        <td style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>H</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style='margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
            <td>
                <h3 class='box-head' style='margin: 0 0 20px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>South Indian chart</h3>
            </td>
        </tr>
        <tr>
            <td>
                <table width='744' style='border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>
                    <tr>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                    </tr>
                    <tr>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;' rowspan='2' colspan='2'>Leo</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                    </tr>
                    <tr>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                    </tr>
                    <tr>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Sun</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>Leo</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>11-33-20</td>
                        <td style='width: 25%;height: 100px;text-align: center;border: 1px solid #ff8c00;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>A</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style='margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
            <td>
                <h3 class='box-head' style='margin: 0 0 20px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>Dasa listing</h3>
            </td>
        </tr>
        <tr>
            <td>
                <table width='744' style='table-layout:fixed;margin:auto;border: 1px solid #ddd;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;'>
                    <thead style='display:table;width:calc(744px - 6px);'>
                        <tr>
                            <th style='border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;	border-bottom: 1px solid #ddd;border-right: 1px solid #ddd;width: 93px;'>A</th>
                            <th style='border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;	border-bottom: 1px solid #ddd;border-right: 1px solid #ddd;width: 93px;'>years</th>
                            <th style='border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;	border-bottom: 1px solid #ddd;border-right: 1px solid #ddd;width: 93px;'>Age group</th>
                            <th style='border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;	border-bottom: 1px solid #ddd;border-right: 1px solid #ddd;width: 93px;'>D</th>
                            <th style='border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;	border-bottom: 1px solid #ddd;border-right: 1px solid #ddd;width: 279px;border-right: 0;'>Planet ruling Natural dasa</th>
                        </tr>
                    </thead>
                    <tbody style='height:150px;overflow-x:hidden;display:block;width:100%;'>
                        <tr style='display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>THe Moon</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>1</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 1 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='    display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>Mars</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>2</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 2 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='    display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>THe Moon</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>1</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 1 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='    display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>Mars</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>2</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 2 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>THe Moon</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>1</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 1 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>Mars</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>2</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 2 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                        <tr style='display:table;width:100%;table-layout:fixed;'>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>THe Moon</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>1</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>upto 1 year</td>
                            <td style='width: 93px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                            <td style='width: 279px;border-collapse: collapse;padding: 8px;color: #33334B;font-family: sans-serif;font-size: 14px;border-right: 1px solid #ddd;'>abc</td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
    </table>
    <table style='margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
            <td>
                <h3 class='original-data' style='margin: 0 0 10px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>Original data provided by the chart owner</h3>
            </td>
        </tr>
        <tr>
            <td>
                <table width='744' style='margin-bottom: 30px;'>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Date of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>20-08-1993</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Place of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>New Delhi</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Time of Birth :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>9:05 AM</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <h3 class='original-data' style='margin-top: 40px;margin: 0 0 10px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>Original data provided by the chart owner</h3></td>
        </tr>
        <tr>
            <td>
                <table width='744'>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Marriage :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>20-08-1993</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>First child :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>New Delhi</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Last call received :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>9:05 AM</p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <p style=' margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 15px;'>Demise of relatives :</p>
                        </td>
                        <td>
                            <p style='margin: 5px 0;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 14px;'>9:05 AM</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style='width:700px;margin: 10px auto 20px;padding: 40px 30px;background:#fff;'>
        <tr>
             <td>
                <h3 style='margin: 0 0 20px;color: #33334B;font-family: sans-serif;font-weight: 600;font-size: 18px;'>North Indian chart</h3>
            </td>
        </tr>
        <tr>
            <td>
                <img src='http://49.50.103.132/dasaImage.png' style='width:100%;' alt='' title=''>
            </td>
        </tr>
    </table>

    
</div>
</body>

</html>"
        End If
        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(mail.Body, Nothing, "text/html")
        mail.AlternateViews.Add(htmlView)
        Dim smtp As New SmtpClient
        smtp.Host = "smtp.gmail.com"
        smtp.Credentials = New NetworkCredential("iamemailsender@gmail.com", "dakiyadaaklaya")
        smtp.EnableSsl = True
        Dim IsMailSent As Boolean = True
        smtp.Send(mail)
    End Sub
End Module
