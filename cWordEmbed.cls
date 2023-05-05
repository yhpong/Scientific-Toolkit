VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWordEmbed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================
'Word Embeddings using either CBOW or Skip-Gram
'All characters are converted to lower case
'note that no one-hot vector is acutally used in the code
'since that would take up a lot of space, words are simply
'coded as integers with corresponding lookup values
'in a dictionary
'=========================================

Private punigram() As Double
Private pnegprob() As Double

Private pSubsampling As Boolean
Private psubprob() As Double

Private pstrMode As String
Private pDict As Scripting.Dictionary
Private pWordCount() As Long

Private pn_word As Long, pn_embed As Long
Private pwin() As Double, pwout() As Double

Private dwin() As Double, dwout() As Double
Private pwin_tmp() As Double, pwout_tmp() As Double
Private pv_avg() As Double, px() As Long

Private pADAM_count As Long
Private dwin1() As Double, dwout1() As Double
Private dwin2() As Double, dwout2() As Double

Private px_progress() As Double

Property Get progress() As Variant
    progress = px_progress
End Property

Property Get win() As Double()
    win = pwin
End Property

Sub PrintEmbed(mysht As Worksheet)
Dim m As Long
    With mysht
        .Range("A1") = pn_word
        .Range("B1") = pn_embed
        .Range("C1") = pstrMode
        .Range("A2").Resize(pn_word, 1).Value = Application.WorksheetFunction.Transpose(pDict.Keys)
        .Range("B2").Resize(pn_word, 1).Value = Application.WorksheetFunction.Transpose(pWordCount)
        m = 2 + pn_word
        .Range("A" & m).Resize(pn_word, pn_embed).Value = pwin
        m = m + pn_word
        .Range("A" & m).Resize(pn_embed, pn_word).Value = pwout
    End With
End Sub

Sub ReadEmbed(mysht As Worksheet)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
    With mysht
        pn_word = .Range("A1")
        pn_embed = .Range("B1")
        pstrMode = .Range("C1")
        ReDim pWordCount(1 To pn_word)
        Set pDict = New Scripting.Dictionary
        For i = 1 To pn_word
            pDict.Add .Range("A" & 1 + i).Value, i
            pWordCount(i) = .Range("B" & 1 + i).Value
        Next i
        m = 2 + pn_word
        ReDim pwin(1 To pn_word, 1 To pn_embed)
        For i = 1 To pn_word
            k = m + i - 1
            For j = 1 To pn_embed
                pwin(i, j) = .Cells(k, j).Value
            Next j
        Next i
        m = m + pn_wor1d
        ReDim pwout(1 To pn_embed, 1 To pn_word)
        For i = 1 To pn_embed
            k = m + i - 1
            For j = 1 To pn_word
                pwout(i, j) = .Cells(k, j).Value
            Next j
        Next i
    End With
End Sub

Sub PrintDict(myRng As Range)
    With myRng
        .Resize(pn_word, 1).Value = Application.WorksheetFunction.Transpose(pDict.Keys)
        .Offset(0, 1).Resize(pn_word, 1).Value = Application.WorksheetFunction.Transpose(pWordCount)
    End With
End Sub


'Build dictionary from input string
Sub BuildDict(strInput As String, Optional isAppend As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim strList() As String
Dim strtmp As String
Dim w() As Double, tmp_x As Double

    If Not isAppend Then
        Set pDict = New Scripting.Dictionary
        ReDim pWordCount(1 To 1)
        n = 0
    Else
        n = pn_word
    End If

    strList = VBA.Split(Trim(LCase(strInput)), " ")

    For i = 0 To UBound(strList, 1)
        strtmp = strList(i)
        If Not pDict.Exists(strtmp) Then
            n = n + 1
            pDict.Add strtmp, n
            ReDim Preserve pWordCount(1 To n)
            pWordCount(n) = 1
        Else
            j = pDict.Item(strtmp)
            pWordCount(j) = pWordCount(j) + 1
        End If
    Next i

    If Not isAppend Then
    
        pn_word = n
        
    Else
    
        m = pn_word
        pn_word = pn_word + n
        
        If ArrayIsEmpty(pwin) Then
            tmp_x = Sqr(2 / pn_embed)
            ReDim Preserve pwout(1 To pn_embed, 1 To pn_word)
            For i = 1 To pn_embed
                For j = m + 1 To pn_word
                    pwout(i, j) = (-0.5 + Rnd()) * tmp_x
                Next j
            Next i
            
            tmp_x = Sqr(2 / pn_word)
            w = pwin
            ReDim pwin(1 To pn_word, 1 To pn_embed)
            For i = 1 To pn_embed
                For j = 1 To m
                    pwin(j, i) = w(j, i)
                Next j
                For j = m + 1 To pn_word
                    pwin(j, i) = (-0.5 + Rnd()) * tmp_x
                Next j
            Next i
        End If
        
    End If
    
End Sub


'convert input string into a vector of integer according to dictionary
Function str2Token(strInput As String) As Long()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim wordIdx() As Long
Dim strWords() As String
    strWords = VBA.Split(strInput, " ")
    
    n = UBound(strWords, 1) + 1
    ReDim wordIdx(1 To n)
    For i = 1 To n
        wordIdx(i) = pDict.Item(strWords(i - 1))
    Next i
    str2Token = wordIdx
End Function

'Main procedure to build the embedding given an input string
'strMode            "CBOW" or "SKIPGRAM"
'n_window           size of sliding window, m = (n_window-1)/2 words before and after the target word
'                   are used as the context. So n_windows must be an odd number >=3.
Sub BuildEmbedding(strMode As String, strInput As String, n_embed As Long, Optional n_window As Long = 5, _
            Optional n_epoch As Long = 10, Optional n_batch As Long = -1, _
            Optional learn_rate As Double = 0.001, _
            Optional useSpeedUp As String = "ADAM", Optional learnSchedule As String = "", _
            Optional err_tol As Double = 0.03, Optional err_tol_rel As Double = 0.001, _
            Optional n_neg As Long = 10, Optional subsampling As Boolean = False, _
            Optional statusShown As Long = 5)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long
Dim wordIdx() As Long
Dim x() As Long, y_tgt() As Long
Dim tmp_x As Double

    If (n_window + 1) Mod 2 <> 0 Or n_window < 3 Then
        Debug.Print "cWordEmbed: BuildEmbedding: n_window needs to be odd number >= 3."
        End
    End If
    
    pstrMode = UCase(Trim(strMode))
    If pstrMode <> "CBOW" And pstrMode <> "SKIPGRAM" And pstrMode <> "SKIPGRAMN" Then
        Debug.Print "cWordEmbed: BuildEmbedding: strMode must be either CBOW, SKIPGRAM or SKIPGRAMN."
        End
    End If
    
    If pstrMode = "SKIPGRAMN" Then
        tmp_x = 0
        ReDim punigram(1 To pn_word)
        ReDim pnegprob(1 To pn_word)
        For i = 1 To pn_word
            punigram(i) = pWordCount(i) / pn_word
            pnegprob(i) = punigram(i) ^ (0.75)
            tmp_x = tmp_x + pnegprob(i)
        Next i
        For i = 1 To pn_word
            pnegprob(i) = pnegprob(i) / tmp_x
        Next i
    End If
    
    'initialze weight matrix
    If ArrayIsEmpty(pwin) Then
        Call Init(n_embed)
    End If
    
    'convert string to integer keys
    wordIdx = str2Token(strInput)
    
    'if subsampling is turned on, pre-calculate the probability to keep each word
    pSubsampling = subsampling
    If pSubsampling Then
        ReDim psubprob(1 To pn_word)
        For i = 1 To pn_word
            tmp_x = pWordCount(i) / pn_word
            psubprob(i) = (Sqr(tmp_x / 0.001) + 1) * 0.001 / tmp_x
        Next i
    End If
    
    'Train network weights
    Call Fit(wordIdx, n_window, n_epoch, n_batch, learn_rate, useSpeedUp, learnSchedule, err_tol, err_tol_rel, n_neg, statusShown)

End Sub


Private Sub Init(n_embed As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
    
    If pn_word = 0 Then
        Debug.Print "cWordEmbed: Init: Use BuildDict to build dictionary first."
        End
    End If
    
    VBA.Randomize
    
    pn_embed = n_embed
    ReDim pwin(1 To pn_word, 1 To pn_embed)
    ReDim pwout(1 To pn_embed, 1 To pn_word)
    
    tmp_x = Sqr(2 / pn_word)
    For j = 1 To pn_embed
        For i = 1 To pn_word
            pwin(i, j) = (-0.5 + Rnd()) * tmp_x
        Next i
    Next j
    
    tmp_x = Sqr(2 / pn_embed)
    For j = 1 To pn_word
        For i = 1 To pn_embed
            pwout(i, j) = (-0.5 + Rnd()) * tmp_x
        Next i
    Next j
    
End Sub


Private Sub createNegSamples(y_neg() As Long, n_neg As Long, y_tgt() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, n_sample As Long
Dim y_idx As Variant
Dim p() As Double, tmp_x As Double

    n_sample = UBound(y_tgt, 2)
    ReDim y_neg(1 To n_neg, 1 To n_sample)
    For iterate = 1 To n_sample
        p = pnegprob
        tmp_x = 1
        For i = 1 To UBound(y_tgt, 1)
            k = y_tgt(i, iterate)
            tmp_x = tmp_x - p(k)
            p(k) = 0
        Next i
        For i = 1 To pn_word
            p(i) = p(i) / tmp_x
        Next i
        y_idx = modMath.Sample(pn_word, n_neg, isReplace:=False, x_prob:=p)
        For i = 1 To n_neg
            y_neg(i, iterate) = y_idx(i)
        Next i
    Next iterate

End Sub


Private Sub createSamples(strMode As String, wordIdx() As Long, n_window As Long, x() As Long, y_tgt() As Long, Optional subsampling As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long
Dim m_offset As Long
Dim wordIdxtmp() As Long

    m_offset = (n_window - 1) / 2
    
    'When subsampling is turn of remove certain words from the universe
    If subsampling Then
        n = 0
        ReDim wordIdxtmp(1 To UBound(wordIdx, 1))
        For i = 1 To UBound(wordIdx, 1)
            If psubprob(wordIdx(i)) > Rnd() Then
                n = n + 1
                wordIdxtmp(n) = wordIdx(i)
            End If
        Next i
        ReDim Preserve wordIdxtmp(1 To n)
    Else
        ReDim pisKeep(1 To pn_word)
        For i = 1 To pn_word
            pisKeep(i) = True
        Next i
        wordIdxtmp = wordIdx
    End If
    
    'create target and context pairs for training
    n_T = UBound(wordIdxtmp)
    
    n = 0
    If strMode = "CBOW" Then
    
        ReDim y_tgt(1 To 1, 1 To n_T)
        ReDim x(1 To n_window - 1, 1 To n_T)
        For iterate = m_offset + 1 To n_T - m_offset
            n = n + 1
            y_tgt(1, n) = wordIdxtmp(iterate)
            For m = 1 To m_offset
                x(m_offset + m, n) = wordIdxtmp(iterate + m)
                x(m_offset - m + 1, n) = wordIdxtmp(iterate - m)
            Next m
        Next iterate
        
        ReDim Preserve y_tgt(1 To 1, 1 To n)
        ReDim Preserve x(1 To n_window - 1, 1 To n)
        
    ElseIf strMode = "SKIPGRAM" Or strMode = "SKIPGRAMN" Then
    
        ReDim y_tgt(1 To n_window - 1, 1 To n_T)
        ReDim x(1 To 1, 1 To n_T)
        For iterate = m_offset + 1 To n_T - m_offset
            n = n + 1
            x(1, n) = wordIdxtmp(iterate)
            For m = 1 To m_offset
                y_tgt(m_offset + m, n) = wordIdxtmp(iterate + m)
                y_tgt(m_offset - m + 1, n) = wordIdxtmp(iterate - m)
            Next m
        Next iterate

        ReDim Preserve x(1 To 1, 1 To n)
        ReDim Preserve y_tgt(1 To n_window - 1, 1 To n)
    
    End If
End Sub


Private Sub ClearHist()
    Erase pv_avg
    Erase px
End Sub


Private Function FwdPass(x() As Long, Optional storeHist As Boolean = False, Optional y_tgt, Optional y_neg) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, iterate As Long
Dim v_avg() As Double, n_window As Long, n_neg As Long
Dim y() As Double
Dim tmp_x As Double, tmp_y As Double
    
    n = UBound(x, 2)
    ReDim y(1 To pn_word, 1 To n)
    ReDim v_avg(1 To pn_embed, 1 To n)
    
    If pstrMode = "CBOW" Then
    
        n_window = UBound(x, 1)
        For iterate = 1 To n
            For j = 1 To n_window
                k = x(j, iterate)
                For i = 1 To pn_embed
                    v_avg(i, iterate) = v_avg(i, iterate) + pwin(k, i)
                Next i
            Next j
            For i = 1 To pn_embed
                v_avg(i, iterate) = v_avg(i, iterate) / n_window
            Next i
        Next iterate
    
    ElseIf pstrMode = "SKIPGRAM" Or pstrMode = "SKIPGRAMN" Then

        For iterate = 1 To n
            k = x(1, iterate)
            For i = 1 To pn_embed
                v_avg(i, iterate) = pwin(k, i)
            Next i
        Next iterate
        
    End If
    
    If pstrMode = "SKIPGRAMN" And Not IsMissing(y_neg) Then
    
        n_neg = UBound(y_neg, 1)
        n_window = UBound(y_tgt, 1)
        ReDim y(1 To n_window + n_neg, 1 To n)
        For iterate = 1 To n
            For i = 1 To n_window + n_neg
            
                If i <= n_window Then
                    k = y_tgt(i, iterate)
                Else
                    k = y_neg(i - n_window, iterate)
                End If
                
                tmp_x = 0
                For j = 1 To pn_embed
                    tmp_x = tmp_x + v_avg(j, iterate) * pwout(j, k)
                Next j
                
                If i <= n_window Then
                    y(i, iterate) = f_sigmoid(tmp_x)
                Else
                    y(i, iterate) = f_sigmoid(-tmp_x)
                End If
                
            Next i
        Next iterate
    
    Else
        
        For iterate = 1 To n
            ReDim y_tmp(1 To pn_word)
            tmp_y = 0
            For i = 1 To pn_word
                tmp_x = 0
                For j = 1 To pn_embed
                    tmp_x = tmp_x + v_avg(j, iterate) * pwout(j, i)
                Next j
                y_tmp(i) = Exp(tmp_x)
                tmp_y = tmp_y + y_tmp(i)
            Next i
        
            For i = 1 To pn_word
                y(i, iterate) = y_tmp(i) / tmp_y
            Next i
        Next iterate
            
    End If
    
    If storeHist Then
        px = x
        pv_avg = v_avg
    End If
    
    FwdPass = y
    
End Function


Private Sub Backward(y, y_tgt() As Long, Optional y_neg As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_window As Long, n_neg As Long, ii As Long
Dim tmp_x As Double, tmp_y As Double
Dim dEdy() As Double, grad_curr() As Double
Dim strtmp As String
    
    n = UBound(px, 2)
    
    If pstrMode = "CBOW" Then
    
        n_window = UBound(px, 1)
        
        ReDim dEdy(1 To pn_word)
        For iterate = 1 To n
        
            k = y_tgt(1, iterate)
            For i = 1 To pn_word
                dEdy(i) = y(i, iterate)
            Next i
            dEdy(k) = dEdy(k) - 1
            
            For j = 1 To pn_word
                For i = 1 To pn_embed
                    dwout(i, j) = dwout(i, j) + dEdy(j) * pv_avg(i, iterate)
                Next i
            Next j
            
            ReDim grad_curr(1 To pn_word)
            For i = 1 To pn_embed
                tmp_x = 0
                For k = 1 To pn_word
                    tmp_x = tmp_x + dEdy(k) * pwout(i, k)
                Next k
                grad_curr(i) = tmp_x / n_window
            Next i
            
            For m = 1 To n_window
                j = px(m, iterate)
                For i = 1 To pn_embed
                    dwin(j, i) = dwin(j, i) + grad_curr(i)
                Next i
            Next m
        Next iterate
    
    ElseIf pstrMode = "SKIPGRAM" Then
    
        n_window = UBound(y_tgt, 1)
        
        For iterate = 1 To n
            
            For j = 1 To pn_word
                For i = 1 To pn_embed
                    dwout(i, j) = dwout(i, j) + n_window * y(j, iterate) * pv_avg(i, iterate)
                Next i
            Next j
            
            For m = 1 To n_window
                k = y_tgt(m, iterate)
                For i = 1 To pn_embed
                    dwout(i, k) = dwout(i, k) - pv_avg(i, iterate)
                Next i
            Next m
            
            ReDim grad_curr(1 To pn_embed)
            For j = 1 To pn_embed
                tmp_x = 0
                For i = 1 To pn_word
                    tmp_x = tmp_x + pwout(j, i) * y(i, iterate)
                Next i
                grad_curr(j) = tmp_x
            Next j

            ReDim dEdy(1 To pn_embed)
            For m = 1 To n_window
                k = y_tgt(m, iterate)
                For j = 1 To pn_embed
                    dEdy(j) = dEdy(j) + pwout(j, k)
                Next j
            Next m
            
            k = px(1, iterate)
            
            For j = 1 To pn_embed
                dwin(k, j) = dwin(k, j) + n_window * grad_curr(j) - dEdy(j)
            Next j
            
        Next iterate
    
    ElseIf pstrMode = "SKIPGRAMN" Then
        
        n_neg = UBound(y_neg, 1)
        n_window = UBound(y_tgt, 1)
        
        For iterate = 1 To n
            
            For m = 1 To n_window + n_neg
                If m <= n_window Then
                    k = y_tgt(m, iterate)
                    For i = 1 To pn_embed
                        dwout(i, k) = dwout(i, k) - pv_avg(i, iterate) * (1 - y(m, iterate))
                    Next i
                Else
                    k = y_neg(m - n_window, iterate)
                    For i = 1 To pn_embed
                        dwout(i, k) = dwout(i, k) + pv_avg(i, iterate) * (1 - y(m, iterate))
                    Next i
                End If
            Next m
            
            ii = px(1, iterate)
            For m = 1 To n_window + n_neg
                If m <= n_window Then
                    k = y_tgt(m, iterate)
                    For j = 1 To pn_embed
                        dwin(ii, j) = dwin(ii, j) - (1 - y(m, iterate)) * pwout(j, k)
                    Next j
                Else
                    k = y_neg(m - n_window, iterate)
                    For j = 1 To pn_embed
                        dwin(ii, j) = dwin(ii, j) + (1 - y(m, iterate)) * pwout(j, k)
                    Next j
                End If
            Next m
            
        Next iterate
    
    End If
    
End Sub


Private Sub Fit(wordIdx() As Long, n_window As Long, Optional n_epoch As Long = 10, Optional n_batch As Long = -1, _
            Optional learn_rate As Double = 0.001, _
            Optional useSpeedUp As String = "ADAM", Optional learnSchedule As String = "", _
            Optional err_tol As Double = 0.3, Optional err_tol_rel As Double = 0.001, _
            Optional n_neg As Long = 10, _
            Optional statusShown As Long = 5)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_T As Long, ii As Long, jj As Long
Dim i_epoch As Long, iterate As Long, n_converge As Long
Dim y() As Double, y_neg() As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_cost As Double, x_cost_prv As Double
Dim strtmp As String
Dim step_size As Double
Dim batchIdx() As Long, batch_size As Long
Dim x() As Long, y_tgt() As Long
Dim x_sub() As Long, y_tgt_sub() As Long
    
    Call createSamples(pstrMode, wordIdx, n_window, x, y_tgt, subsampling:=False)
    n = UBound(x, 2)
    batch_size = Int(n / n_batch)
    
    If pstrMode = "SKIPGRAMN" Then
        Call createNegSamples(y_neg, n_neg, y_tgt)
        y = FwdPass(x, storeHist:=False, y_tgt:=y_tgt, y_neg:=y_neg)
    Else
        y = FwdPass(x, storeHist:=False)
    End If
    x_cost_prv = calcLoss(y, y_tgt)
    
    n_converge = 0
    step_size = learn_rate
    ReDim px_progress(1 To 3, 1 To 1)
    For i_epoch = 1 To n_epoch
        
        DoEvents
        If (i_epoch - 1) Mod statusShown = 0 Then
            Application.StatusBar = "cWordEmbed: Fit: " & i_epoch & "/" & n_epoch & "..."
        End If
        
        Call CacheCurrentWgt
        
        'if subsampling is turned on, recreate a new set of samples every epoch
        If pSubsampling Then
            Call createSamples(pstrMode, wordIdx, n_window, x, y_tgt, subsampling:=True)
            n = UBound(x, 2)
            batch_size = Int(n / n_batch)
        End If
        
        If n_batch <= 1 Then
            
            Call InitWgtChg
            If pstrMode = "SKIPGRAMN" Then
                Call createNegSamples(y_neg, n_neg, y_tgt)
                y = FwdPass(x, storeHist:=True, y_tgt:=y_tgt, y_neg:=y_neg)
                Call Backward(y, y_tgt, y_neg)
            Else
                y = FwdPass(x, storeHist:=True)
                Call Backward(y, y_tgt)
            End If
            Call UpdateWgt(step_size, useSpeedUp:=useSpeedUp)
            Call ClearHist
            
        Else
            batchIdx = Shuffle(n)
            
            ii = 0: jj = 0
            Do While (jj + 1) <= n

                'Extract a mini-batch
                ii = jj + 1
                jj = jj + batch_size
                If jj > n Then jj = n
                m = jj - ii + 1
                If pstrMode = "CBOW" Then
                    ReDim x_sub(1 To n_window - 1, 1 To m)
                    ReDim y_tgt_sub(1 To 1, 1 To m)
                    For i = 1 To m
                        k = ii + i - 1
                        y_tgt_sub(1, i) = y_tgt(1, batchIdx(k))
                        For j = 1 To n_window - 1
                            x_sub(j, i) = x(j, batchIdx(k))
                        Next j
                    Next i
                ElseIf pstrMode = "SKIPGRAM" Or pstrMode = "SKIPGRAMN" Then
                    ReDim x_sub(1 To 1, 1 To m)
                    ReDim y_tgt_sub(1 To n_window - 1, 1 To m)
                    For i = 1 To m
                        k = ii + i - 1
                        x_sub(1, i) = x(1, batchIdx(k))
                        For j = 1 To n_window - 1
                            y_tgt_sub(j, i) = y_tgt(j, batchIdx(k))
                        Next j
                    Next i
                End If
                
                Call InitWgtChg
                If pstrMode = "SKIPGRAMN" Then
                    Call createNegSamples(y_neg, n_neg, y_tgt_sub)
                    y = FwdPass(x_sub, storeHist:=True, y_tgt:=y_tgt_sub, y_neg:=y_neg)
                    Call Backward(y, y_tgt_sub, y_neg)
                Else
                    y = FwdPass(x_sub, storeHist:=True)
                    Call Backward(y, y_tgt_sub)
                End If
                Call UpdateWgt(step_size, useSpeedUp:=useSpeedUp)
                Call ClearHist

            Loop
        
        End If
        
        'if subsampling is turned on, use full sample to evaluate performance
        'maybe faster to save full sample in the beginning, but recreat it every epoch saves some memory
        If pSubsampling Then
            Call createSamples(pstrMode, wordIdx, n_window, x, y_tgt, subsampling:=False)
            n = UBound(x, 2)
            batch_size = Int(n / n_batch)
        End If
        
        If pstrMode = "SKIPGRAMN" Then
            Call createNegSamples(y_neg, n_neg, y_tgt)
            y = FwdPass(x, storeHist:=False, y_tgt:=y_tgt, y_neg:=y_neg)
        Else
            y = FwdPass(x, storeHist:=False)
        End If
        x_cost = calcLoss(y, y_tgt)
        
        ReDim Preserve px_progress(1 To 3, 1 To i_epoch)
        px_progress(1, i_epoch) = i_epoch
        px_progress(2, i_epoch) = x_cost
        px_progress(3, i_epoch) = step_size
        DoEvents
        Debug.Print "Epoch " & i_epoch & "/" & n_epoch & ", cost=" & Format(x_cost, "0.0000E+00") & ", step_size=" & Format(step_size, "0.0000E+00")
        
        If (x_cost <= x_cost_prv) And (x_cost < err_tol Or Abs(x_cost_prv - x_cost) <= Abs(err_tol_rel * x_cost_prv)) Then
            n_converge = n_converge + 1
        Else
            n_converge = 0
        End If
        If n_converge >= 5 Then Exit For
        
        If learnSchedule = "AGGRESSIVE" Then
            If (x_cost <= x_cost_prv) Then
                step_size = step_size * 1.05
                x_cost_prv = x_cost
            Else
                step_size = step_size * 0.1
                Call RestoreWgt
                DoEvents
                Debug.Print "cost increases, dicard current epoch."
                If UCase(useSpeedUp) = "ADAM" Then Call ADAM_Init
                If step_size < 0.000000001 Then Exit For
            End If
        ElseIf learnSchedule = "DECAY" Then
            step_size = learn_rate * (1 - i_epoch / n_epoch)
            x_cost_prv = x_cost
        Else
            x_cost_prv = x_cost
        End If

    Next i_epoch
    
    If UCase(useSpeedUp) = "ADAM" Then Call ADAM_Clear
    Call ClearCacheWgt
    Call ClearWgtChg
    Application.StatusBar = False
End Sub


Private Function calcLoss(y() As Double, y_tgt() As Long) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, mm As Long
Dim tmp_x As Double

    tmp_x = 0
    n = UBound(y, 2)
    
    If pstrMode = "CBOW" Then

        For iterate = 1 To n
            k = y_tgt(1, iterate)
            tmp_x = tmp_x - Log(y(k, iterate))
        Next iterate
    
    ElseIf pstrMode = "SKIPGRAM" Then
    
        m = UBound(y_tgt, 1)
        For iterate = 1 To n
            For j = 1 To m
                k = y_tgt(j, iterate)
                tmp_x = tmp_x - Log(y(k, iterate))
            Next j
        Next iterate
        
        tmp_x = tmp_x / m
        
    ElseIf pstrMode = "SKIPGRAMN" Then
    
        m = UBound(y, 1)
        For iterate = 1 To n
            For j = 1 To m
                tmp_x = tmp_x + Log(y(m, iterate))
            Next j
        Next iterate
        tmp_x = -tmp_x / m
        
    End If
    
    calcLoss = tmp_x / n
    
End Function

'Clear and reset memories used in ADAM
Private Sub ADAM_Clear()
    pADAM_count = 0
    Erase dwout1, dwin1
    Erase dwout2, dwin2
End Sub

Private Sub ADAM_Init()
    pADAM_count = 0
    ReDim dwin1(1 To pn_word, 1 To pn_embed)
    ReDim dwout1(1 To pn_embed, 1 To pn_word)
    ReDim dwin2(1 To pn_word, 1 To pn_embed)
    ReDim dwout2(1 To pn_embed, 1 To pn_word)
End Sub

Private Sub ADAM_CalcMoment()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
    
    If pADAM_count = 0 Then Call ADAM_Init
    
    pADAM_count = pADAM_count + 1
    
    For j = 1 To pn_embed
        For i = 1 To pn_word
            dwin1(i, j) = 0.9 * dwin1(i, j) + 0.1 * dwin(i, j)
            dwin2(i, j) = 0.999 * dwin2(i, j) + 0.001 * dwin(i, j) ^ 2
        Next i
    Next j
    
    For i = 1 To pn_word
        For j = 1 To pn_embed
            dwout1(j, i) = 0.9 * dwout1(j, i) + 0.1 * dwout(j, i)
            dwout2(j, i) = 0.999 * dwout2(j, i) + 0.001 * dwout(j, i) ^ 2
        Next j
    Next i

End Sub

'Cache current weights
Private Sub CacheCurrentWgt()
    pwin_tmp = pwin
    pwout_tmp = pwout
End Sub

'Clear cache weights
Private Sub ClearCacheWgt()
    Erase pwin_tmp, pwout_tmp
End Sub

'Restore weights to cached values
Private Sub RestoreWgt()
    pwin = pwin_tmp
    pwout = pwout_tmp
End Sub

'initialize gradients
Private Sub InitWgtChg()
    ReDim dwin(1 To pn_word, 1 To pn_embed)
    ReDim dwout(1 To pn_embed, 1 To pn_word)
End Sub

'Clear all gradients
Private Sub ClearWgtChg()
    Erase dwin, dwout
End Sub


'Applied stored gradients to update weights
'stored gradients are erased once used
Private Sub UpdateWgt(learn_rate As Double, Optional useSpeedUp As String = "")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim iterate As Long
Dim adam_discount1 As Double, adam_discount2 As Double

    If UCase(useSpeedUp) = "ADAM" Then
        
        Call ADAM_CalcMoment
        
        adam_discount1 = 1 / (1 - 0.9 ^ pADAM_count)
        adam_discount2 = 1 / (1 - 0.999 ^ pADAM_count)
        
        For j = 1 To pn_word
            For i = 1 To pn_embed
                pwout(i, j) = pwout(i, j) - learn_rate * dwout1(i, j) * adam_discount1 / (Sqr(dwout2(i, j) * adam_discount2) + 0.00000001)
            Next i
        Next j
        
        For i = 1 To pn_embed
            For j = 1 To pn_word
                pwin(j, i) = pwin(j, i) - learn_rate * dwin1(j, i) * adam_discount1 / (Sqr(dwin2(j, i) * adam_discount2) + 0.00000001)
            Next j
        Next i
        
    Else
        
        For j = 1 To pn_word
            For i = 1 To pn_embed
                pwout(i, j) = pwout(i, j) - learn_rate * dwout(i, j)
            Next i
        Next j
        
        For i = 1 To pn_embed
            For j = 1 To pn_word
                pwin(j, i) = pwin(j, i) - learn_rate * dwin(j, i)
            Next j
        Next i

    End If
    Erase dwin, dwout
    
End Sub



'Generate a randomly order seqeunce from 1:n
Private Function Shuffle(n As Long) As Long()
Dim i As Long, j As Long
Dim k As Long
Dim x() As Long
Dim vtmp As Variant
    ReDim x(1 To n)
    For i = 1 To n
        x(i) = i
    Next i
    Randomize
    For i = n To 2 Step -1
        j = Int(Rnd() * i) + 1  'Random_Integer(1, i)
        vtmp = x(j)
        x(j) = x(i)
        x(i) = vtmp
    Next i
    Shuffle = x
End Function


Private Function ArrayIsEmpty(x) As Boolean
Dim i As Long
    If Not IsArray(x) Then
        ArrayIsEmpty = True
    Else
        ArrayIsEmpty = False
        On Error Resume Next
        i = UBound(x, 1)
        If Err.Number <> 0 Then
            Err.Clear
            ArrayIsEmpty = True
        End If
    End If
End Function


Private Function f_sigmoid(x As Double)
    If x > 20 Then
        f_sigmoid = 1
    ElseIf x < -20 Then
        f_sigmoid = 0
    Else
        f_sigmoid = 1 / (1 + Exp(-x))
    End If
End Function


