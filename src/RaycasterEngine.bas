Attribute VB_Name = "Module1"
Option Explicit

' --- 1. WINDOWS API DECLARATION ---
#If VBA7 Then
    Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' --- 2. THE RAYCASTER ENGINE ---
Sub StartRaycaster()


    ' The World Map (16x16 String)
    Dim sMap As String
    sMap = sMap & "################"
    sMap = sMap & "#..............#"
    sMap = sMap & "#.......########"
    sMap = sMap & "#..............#"
    sMap = sMap & "#......##......#"
    sMap = sMap & "#......##......#"
    sMap = sMap & "#..............#"
    sMap = sMap & "###............#"
    sMap = sMap & "##.............#"
    sMap = sMap & "#......####..###"
    sMap = sMap & "#......#.......#"
    sMap = sMap & "#......#.......#"
    sMap = sMap & "#..............#"
    sMap = sMap & "#......#########"
    sMap = sMap & "#..............#"
    sMap = sMap & "################"
    
    
    ' Screen and Map Settings
    Dim nScreenWidth As Integer: nScreenWidth = 120
    Dim nScreenHeight As Integer: nScreenHeight = 40
    Dim nMapWidth As Integer: nMapWidth = 16
    Dim nMapHeight As Integer: nMapHeight = 16
    
    ' Player Variables
    Dim fPlayerX As Single: fPlayerX = 8#
    Dim fPlayerY As Single: fPlayerY = 8#
    Dim fPlayerA As Single: fPlayerA = 0#
    Dim fFOV As Single: fFOV = 3.14159 / 4#
    Dim fDepth As Single: fDepth = 16#
    
' --- Speed & Movement Settings ---
    Dim fMoveSpeed As Single: fMoveSpeed = Val(Range("DT44").Value) / 10#
    Dim fRotSpeed As Single: fRotSpeed = Val(Range("DT45").Value) / 10#
    
    ' Defaulting to 1 and 2 if the cells are empty
    If fMoveSpeed <= 0 Then fMoveSpeed = 0.2
    If fRotSpeed <= 0 Then fRotSpeed = 0.2
    
    ' --- Health & Combat Stats ---
    Dim fMaxHealth As Single: fMaxHealth = Val(Range("DT46").Value)
    If fMaxHealth <= 0 Then fMaxHealth = 100
    
    Dim fPlayerHealth As Single: fPlayerHealth = fMaxHealth
    
    ' Scaled Damage: UI 10 = 0.1 damage per frame
    Dim fDamageValue As Single: fDamageValue = Val(Range("DT47").Value) / 100
    If fDamageValue <= 0 Then fDamageValue = 0.05
    
    Dim nKillCount As Long: nKillCount = 0
    
' --- Enemy Variables ---
    Dim nMaxEnemies As Integer: nMaxEnemies = Val(Range("DT43").Value)
    If nMaxEnemies < 1 Then nMaxEnemies = 1
    
    Dim fEX() As Single, fEY() As Single, bEAlive() As Boolean, fETimer() As Single
    ReDim fEX(1 To nMaxEnemies), fEY(1 To nMaxEnemies), bEAlive(1 To nMaxEnemies), fETimer(1 To nMaxEnemies)
    
    Dim iE As Integer
    Randomize
    
    ' Initial Random Placement
    For iE = 1 To nMaxEnemies
        Dim bPlaced As Boolean: bPlaced = False
        Do While Not bPlaced
            Dim rx As Integer: rx = Int(Rnd * 14) + 1
            Dim ry As Integer: ry = Int(Rnd * 14) + 1
            ' Only place if it's a floor (.) and away from player start
            If Mid(sMap, rx * 16 + ry + 1, 1) = "." Then
                If Abs(rx - fPlayerX) > 2 Or Abs(ry - fPlayerY) > 2 Then
                    fEX(iE) = rx + 0.5: fEY(iE) = ry + 0.5
                    bEAlive(iE) = True
                    fETimer(iE) = 0
                    bPlaced = True
                End If
            End If
        Loop
    Next iE
    
    
    ' Projectile Variables
    Dim fProjX As Single, fProjY As Single
    Dim fProjDX As Single, fProjDY As Single
    Dim bProjActive As Boolean: bProjActive = False
    
    

    Dim screenArray() As Variant
    ReDim screenArray(1 To nScreenHeight, 1 To nScreenWidth)
    
       
    Do
        DoEvents
        
' --- 1. INPUT HANDLING ---
        ' Rotation
        If GetAsyncKeyState(vbKeyA) <> 0 Then fPlayerA = fPlayerA - fRotSpeed
        If GetAsyncKeyState(vbKeyD) <> 0 Then fPlayerA = fPlayerA + fRotSpeed
        
        ' Forward / Backward
        If GetAsyncKeyState(vbKeyW) <> 0 Then
            fPlayerX = fPlayerX + Sin(fPlayerA) * fMoveSpeed
            fPlayerY = fPlayerY + Cos(fPlayerA) * fMoveSpeed
            ' Collision check
            If Mid(sMap, Int(fPlayerX) * 16 + Int(fPlayerY) + 1, 1) = "#" Then
                fPlayerX = fPlayerX - Sin(fPlayerA) * fMoveSpeed
                fPlayerY = fPlayerY - Cos(fPlayerA) * fMoveSpeed
            End If
        End If
        
        If GetAsyncKeyState(vbKeyS) <> 0 Then
            fPlayerX = fPlayerX - Sin(fPlayerA) * fMoveSpeed
            fPlayerY = fPlayerY - Cos(fPlayerA) * fMoveSpeed
            ' Collision check
            If Mid(sMap, Int(fPlayerX) * 16 + Int(fPlayerY) + 1, 1) = "#" Then
                fPlayerX = fPlayerX + Sin(fPlayerA) * fMoveSpeed
                fPlayerY = fPlayerY + Cos(fPlayerA) * fMoveSpeed
            End If
        End If
        
        ' FIRE
        If GetAsyncKeyState(vbKeySpace) <> 0 And Not bProjActive Then
            bProjActive = True
            fProjX = fPlayerX: fProjY = fPlayerY
            fProjDX = Sin(fPlayerA) * 0.5: fProjDY = Cos(fPlayerA) * 0.5
        End If
        
        If GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Do

        ' --- 2. PROJECTILE PHYSICS ---
        If bProjActive Then
            fProjX = fProjX + fProjDX
            fProjY = fProjY + fProjDY
            If Mid(sMap, Int(fProjX) * nMapWidth + Int(fProjY) + 1, 1) = "#" Then bProjActive = False
            If fProjX < 0 Or fProjX >= nMapWidth Or fProjY < 0 Or fProjY >= nMapHeight Then bProjActive = False
        End If
        
' --- 2b. ENEMY AI & SPAWN SYSTEM ---
        Dim nThresh As Long: nThresh = Val(Range("DT42").Value)
        If nThresh <= 0 Then nThresh = 100

        For iE = 1 To nMaxEnemies
            If bEAlive(iE) Then
                ' AI Movement
                fEX(iE) = fEX(iE) + Sgn(fPlayerX - fEX(iE)) * 0.02
                fEY(iE) = fEY(iE) + Sgn(fPlayerY - fEY(iE)) * 0.02
                
' --- PLAYER DAMAGE LOGIC ---
        ' Check distance between this specific enemy and the player
        Dim fDistToPlayer As Single
        fDistToPlayer = Sqr((fEX(iE) - fPlayerX) ^ 2 + (fEY(iE) - fPlayerY) ^ 2)
        
        If fDistToPlayer < 0.6 Then
            fPlayerHealth = fPlayerHealth - fDamageValue
        End If
        
                ' Collision with Projectile
                If bProjActive Then
                    If Sqr((fProjX - fEX(iE)) ^ 2 + (fProjY - fEY(iE)) ^ 2) < 0.7 Then
                        bEAlive(iE) = False
                        bProjActive = False
                        nKillCount = nKillCount + 1
                        fETimer(iE) = 0 ' Start this specific enemy's cooldown
                    End If
                End If
            Else
                ' Individual Respawn Timer for this index
                fETimer(iE) = fETimer(iE) + 1
                
                If fETimer(iE) > nThresh Then
                    Dim nTX As Integer: nTX = Int(Rnd * 14) + 1
                    Dim nTY As Integer: nTY = Int(Rnd * 14) + 1
                    ' Only spawn on floor and away from player
                    If Mid(sMap, nTX * 16 + nTY + 1, 1) = "." Then
                        If Sqr((nTX - fPlayerX) ^ 2 + (nTY - fPlayerY) ^ 2) > 4 Then
                            fEX(iE) = nTX + 0.5: fEY(iE) = nTY + 0.5
                            bEAlive(iE) = True
                            fETimer(iE) = 0
                        End If
                    End If
                End If
            End If
        Next iE

        ' --- 3. 3D RENDERING ---
        Dim x As Integer, y As Integer
        
        ' Pre-calculate Projectile & Enemy screen positions once per frame
        Dim fAngleToProj As Single, fDistToProj As Single, fProjScreenX As Integer
        Dim fAngleToEnemy As Single, fDistToEnemy As Single, fEnemyScreenX As Integer
        
        If bProjActive Then
            fDistToProj = Sqr((fProjX - fPlayerX) ^ 2 + (fProjY - fPlayerY) ^ 2)
            fAngleToProj = Atn2(fProjX - fPlayerX, fProjY - fPlayerY)
            Dim fPDiff As Single: fPDiff = fAngleToProj - fPlayerA
            Do While fPDiff < -3.14159: fPDiff = fPDiff + 6.28318: Loop
            Do While fPDiff > 3.14159: fPDiff = fPDiff - 6.28318: Loop
            fProjScreenX = Int((fPDiff / fFOV + 0.5) * nScreenWidth)
        End If
        

        For x = 0 To nScreenWidth - 1
            Dim fRayAngle As Single: fRayAngle = (fPlayerA - fFOV / 2#) + (x / nScreenWidth) * fFOV
            Dim fDistanceToWall As Single: fDistanceToWall = 0#
            Dim bHitWall As Boolean: bHitWall = False
            
            Dim fEyeX As Single: fEyeX = Sin(fRayAngle)
            Dim fEyeY As Single: fEyeY = Cos(fRayAngle)
            
            Do While Not bHitWall And fDistanceToWall < fDepth
                fDistanceToWall = fDistanceToWall + 0.1
                Dim nTestX As Integer: nTestX = Int(fPlayerX + fEyeX * fDistanceToWall)
                Dim nTestY As Integer: nTestY = Int(fPlayerY + fEyeY * fDistanceToWall)
                
                If nTestX < 0 Or nTestX >= nMapWidth Or nTestY < 0 Or nTestY >= nMapHeight Then
                    bHitWall = True: fDistanceToWall = fDepth
                Else
                    If Mid(sMap, nTestX * nMapWidth + nTestY + 1, 1) = "#" Then bHitWall = True
                End If
            Loop
            
            Dim nCeiling As Integer: nCeiling = (nScreenHeight / 2#) - nScreenHeight / fDistanceToWall
            Dim nFloor As Integer: nFloor = nScreenHeight - nCeiling
            
           ' Shading based on distance
            Dim nShade As String
            If fDistanceToWall <= fDepth / 4# Then
                nShade = "@"
            ElseIf fDistanceToWall < fDepth / 3# Then
                nShade = "#"
            ElseIf fDistanceToWall < fDepth / 2# Then
                nShade = "x"
            ElseIf fDistanceToWall < fDepth Then
                nShade = "."
            Else
                nShade = " "
            End If
            ' Draw Wall, Sky, Floor
            For y = 1 To nScreenHeight
                If y <= nCeiling Then
                    screenArray(y, x + 1) = " "
                ElseIf y > nCeiling And y <= nFloor Then
                    screenArray(y, x + 1) = nShade
                Else
                    screenArray(y, x + 1) = "."
                End If
            Next y
            
' --- 3b. OVERLAY ALL ENEMIES  ---
            For iE = 1 To nMaxEnemies
                If bEAlive(iE) Then
                    ' Calculate distance and angle for THIS specific enemy
                    Dim fVEX As Single: fVEX = fEX(iE) - fPlayerX
                    Dim fVEY As Single: fVEY = fEY(iE) - fPlayerY
                    Dim fDistE As Single: fDistE = Sqr(fVEX * fVEX + fVEY * fVEY)
                    
                    Dim fAngE As Single: fAngE = Atn2(fVEX, fVEY)
                    Dim fDiffE As Single: fDiffE = fAngE - fPlayerA
                    
                    ' Normalize angle
                    Do While fDiffE < -3.14159: fDiffE = fDiffE + 6.28318: Loop
                    Do While fDiffE > 3.14159: fDiffE = fDiffE - 6.28318: Loop
                    
                    ' Project to screen
                    Dim nEScreenX As Integer
                    nEScreenX = Int((fDiffE / fFOV + 0.5) * nScreenWidth)
                    
                    ' Dynamic Size & Thickness
                    Dim nEHz As Integer: nEHz = nScreenHeight / fDistE
                    Dim nEHalfW As Integer: nEHalfW = 2 + Int(nEHz / 10) ' Gets thicker as it approaches
                    
                    ' Draw if the ray is currently in the enemy's horizontal "hitbox"
                    If x >= (nEScreenX - nEHalfW) And x <= (nEScreenX + nEHalfW) Then
                        ' Depth Check: Only draw if the enemy is closer than the wall
                        If fDistE < fDistanceToWall Then
                            For y = (nScreenHeight / 2) - (nEHz / 4) To (nScreenHeight / 2) + (nEHz / 4)
                                If y >= 1 And y <= nScreenHeight Then screenArray(y, x + 1) = "M"
                            Next y
                        End If
                    End If
                End If
            Next iE
            
' OVERLAY PROJECTILE
            If bProjActive Then
                ' 1. Calculate Dynamic Radius (Same as vertical logic)
                Dim nProjSize As Integer: nProjSize = nScreenHeight / fDistToProj
                Dim nRadius As Integer: nRadius = nProjSize / 10
                If nRadius < 1 Then nRadius = 1 ' Ensure it doesn't disappear at distance
                
                ' 2. Check if current column 'x' is within the horizontal arms
                If x >= (fProjScreenX - nRadius) And x <= (fProjScreenX + nRadius) Then
                    If fDistToProj < fDistanceToWall Then
                        Dim nMidY As Integer: nMidY = nScreenHeight / 2
                        
                        ' Vertical Beam (Only in the center column)
                        If x = fProjScreenX Then
                            For y = nMidY - nRadius To nMidY + nRadius
                                If y >= 1 And y <= nScreenHeight Then screenArray(y, x + 1) = "|"
                            Next y
                            ' Center junction
                            screenArray(nMidY, x + 1) = "+"
                        Else
                            ' Horizontal Arms (Only at the midpoint)
                            If nMidY >= 1 And nMidY <= nScreenHeight Then
                                screenArray(nMidY, x + 1) = "-"
                            End If
                        End If
                    End If
                End If
            End If
            
            
        Next x
        
        ' --- 3c. GUN OVERLAY (HUD) ---
        ' This draws a static barrel at the bottom center of the screen (Column 60)
        Dim gX As Integer, gY As Integer
        For gY = 33 To 40 ' Bottom 8 rows
            For gX = 55 To 65 ' Center of 120 columns
                ' Simple Gun Shape
                If gY >= 37 Then
                    ' Base/Body of the weapon
                    If gX >= 57 And gX <= 63 Then screenArray(gY, gX) = "W"
                Else
                    ' The Barrel
                    If gX >= 60 And gX <= 60 Then screenArray(gY, gX) = "!"
                End If
            Next gX
        Next gY
        

' --- 4. MINIMAP (With Directional Arrow) ---
        Dim nx As Integer, ny As Integer
        Dim pIcon As String
        Dim fTempA As Single: fTempA = fPlayerA
        
        ' Normalize angle to 0 - 6.28 (2*PI) range
        Do While fTempA < 0: fTempA = fTempA + 6.28318: Loop
        Do While fTempA >= 6.28318: fTempA = fTempA - 6.28318: Loop
        
        ' Determine arrow based on angle (North = 0)
        If fTempA >= 5.497 Or fTempA < 0.785 Then
            pIcon = "^" ' Looking Up
        ElseIf fTempA >= 0.785 And fTempA < 2.356 Then
            pIcon = ">" ' Looking Right
        ElseIf fTempA >= 2.356 And fTempA < 3.927 Then
            pIcon = "v" ' Looking Down
        Else
            pIcon = "<" ' Looking Left
        End If

        For nx = 0 To nMapWidth - 1
            For ny = 0 To nMapHeight - 1
                Dim mapChar As String
                mapChar = Mid(sMap, nx * nMapWidth + ny + 1, 1)
                
                ' Draw Player Arrow
                If nx = Int(fPlayerX) And ny = Int(fPlayerY) Then
                    mapChar = pIcon
                End If
                
                ' Draw Enemy
                    ' Draw Enemies on Minimap
                For iE = 1 To nMaxEnemies
                    If bEAlive(iE) And nx = Int(fEX(iE)) And ny = Int(fEY(iE)) Then
                        mapChar = "E"
                    End If
                Next iE
                
                ' Draw Projectile
                If bProjActive And nx = Int(fProjX) And ny = Int(fProjY) Then
                    mapChar = "|"
                End If
                
                screenArray(nx + 2, ny + 2) = mapChar
            Next ny
        Next nx
        
        ' Update Health Display in cell DS46
        Range("DT52").Value = "" & Int(fPlayerHealth)
        Range("DT53").Value = nKillCount
        
        ' GAME OVER CHECK
        If fPlayerHealth <= 0 Then
            Range("DT53").Value = nKillCount
            
            MsgBox "GAME OVER!" & vbCrLf & _
            "Kill Count: " & nKillCount
            Exit Do
        End If
        
        Range("A1").Resize(nScreenHeight, nScreenWidth).Value = screenArray
    Loop
End Sub

' --- 5. ATAN2 HELPER (Corrected Block Syntax) ---
Function Atn2(ByVal x As Single, ByVal y As Single) As Single
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atn2 = Atn(y / x) + 3.14159
        Else
            Atn2 = Atn(y / x) - 3.14159
        End If
    Else
        ' Case where x = 0 to avoid division by zero
        If y > 0 Then
            Atn2 = 1.57079
        ElseIf y < 0 Then
            Atn2 = -1.57079
        Else
            Atn2 = 0
        End If
    End If
    
    ' Rotate 90 degrees to align with your North=0 system
    Atn2 = 1.57079 - Atn2
End Function

