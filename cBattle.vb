''' <summary>Класс, управляющий битвой </summary>
Public Class cBattle
    Public Enum IsEnemyEnum As Byte
        Enemy = 0
        Friendly = 1
        Neutral = 2
        CheckError = 255
    End Enum

    Public Enum PropertyVisualizationStyleEnum As Byte
        HORIZONTAL_TOP = 0
        HORIZONTAL_BOTTOM = 1
        VERTICAL_LEFT = 2
        VERTICAL_RIGHT = 3
        VERTICAL_FRONT = 4
        VERTICAL_BACK = 5
    End Enum

    Public Enum BattleFinishEnum As Byte
        CUSTOM = 0
        ENEMIES_DEAD = 1
        MAIN_HERO_DEAD = 2
        ALL_IS_DEAD = 3
        CHECK_ERROR = 255
    End Enum

    Public Enum MagicTypesEnum As Byte
        Damage = 0
        TotalDamage = 1
        SelfEnhancer = 2
        FriendsEnhancer = 3
        Summoning = 4
        OtherWithvictim = 5
        OtherWithoutVictim = 6
        OtherFriend = 7
    End Enum

    Private Enum victimChoiceEnum As Byte
        MainHeroOnly = 0
        NotMainHero = 1
        Randomly = 2
        TheStrongest = 3
        TheWeakest = 4
    End Enum

    ''' <summary>Класс для сортировки бойцов по расположению на поле боя</summary>
    Public Class FightersPosition
        Public armyId As Integer
        Public fromLeft As Boolean
        ''' <summary>ключ - ранг юнита, значение - список бойцов данного ранга</summary>
        Public orders As SortedList(Of Integer, List(Of Integer))
    End Class

    ''' <summary>Класс для подбора бойца, которых должен пойти следующим</summary>
    Private Class CandidatesToBeNext
        Public fighterId As Integer
        Public speed As Double
        Public turns As Integer
    End Class

    ''' <summary>Класс для хранения информации о бойцах</summary>
    Public Class cFighters
        ''' <summary>свойства персонажа, ассоциированного с данным бойцом</summary>
        Public heroProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType)
        ''' <summary>Id персонажа, ассоциированного с данным бойцом</summary>
        Public heroId As Integer
        ''' <summary>Id книги магий песонажа до дублирования</summary>
        Public magicBookId As Integer = -1
        ''' <summary>Id набора способностей песонажа до дублирования</summary>
        Public abilitySetId As Integer = -1
        ''' <summary>Является ли боец дубликатом персонажа (или же он первый/в единственном числе)</summary>
        Public isDuplicate As Boolean = False
        ''' <summary>Id армии бойца</summary>
        Public armyId As Integer = -1
        ''' <summary>Id юнита в армии</summary>
        Public armyUnitId As Integer = -1
    End Class

    ''' <summary>Главное хранилище информации о бойцах</summary>
    Public Fighters As New List(Of cFighters)
    ''' <summary>Список с Id бойцов из Fighters, доступных для атаки / воздействия</summary>
    Public lstVictims As New List(Of Integer)
    ''' <summary>Список с информацией для распределения бойцов по позициям на поле боя</summary>
    Public lstPositions As New List(Of FightersPosition)
    ''' <summary>Хранение книг магий для возможности восстановления (ключ - Id героя (на бойца), значение - книга магий</summary>
    Public lstMagicBooksCopies As New SortedList(Of Integer, SortedList(Of String, MatewScript.ChildPropertiesInfoType))
    ''' <summary>Хранение способностей для возможности восстановления (ключ - Id героя (на бойца), значение - набор способностей</summary>
    Public lstAbilityCopies As New SortedList(Of Integer, SortedList(Of String, MatewScript.ChildPropertiesInfoType))
    ''' <summary>Обозначает что переход на новую локацию совершен сразу после боя (надо перезагрузить html-документ)</summary>
    Public JustAfterBattle As Boolean = False
    ''' <summary>Для получания позиции атакующего бойца</summary>
    Private FighterFromLeft As Boolean = False
    ''' <summary>Список аудио, который играл до боя</summary>
    Private audioListBeforeBattle As Integer = -1
    ''' <summary>Таймер, служащий для перемещения бойцов, указанных в lsthFightersToMoveInto и lstFightersToMoveOut</summary>
    Public WithEvents timMoveFighters As New Timer With {.Enabled = False, .Interval = 100}
    ''' <summary>История боя</summary>
    Public lstHistory As New List(Of String)
    ''' <summary>True если в данный момент происходит уничтожение призванных бойцов вслед за хозяином</summary>
    Public killingSummoned As Boolean = False

    Dim lsthFightersToMoveInto As New List(Of HtmlElement)
    Public lstFightersToMoveOut As New List(Of Integer)

#Region "Fighters management"
    ''' <summary>
    ''' Заявляет бойца на бой
    ''' </summary>
    ''' <param name="heroId">Id персонажа, ассоциированного с данным героем</param>
    ''' <param name="fName">Имя героя в структуре Fighters</param>
    ''' <param name="armyId">Id армии, которой принадлежит персонаж</param>
    ''' <returns>Id бойца в Fighters</returns>
    Public Function AddFighter(ByVal heroId As Integer, ByVal fName As String, Optional ByVal armyId As Integer = -1, Optional ByVal armyUnitId As Integer = -1, Optional owner As Integer = -1) As Integer
        'Дублировать книгу заклинаний и наборы способностей персонажей, если на бой заявлены их копии

        Dim classH As Integer = mScript.mainClassHash("H"), classMg As Integer = mScript.mainClassHash("Mg"), classAb As Integer = mScript.mainClassHash("Ab"), arrHeroId() As String = {heroId.ToString}
        If String.IsNullOrEmpty(UnWrapString(fName)) Then
            fName = GetNewName(mScript.mainClass(classH).ChildProperties(heroId)("Name").Value)
        Else
            Dim fId As Integer = GetFighterByName(fName)
            If fId > -1 Then
                mScript.LAST_ERROR = "Герой с именем " + Chr(34) + fName + Chr(34) + " уже учавствует в битве! Укажите в этой функции его новое имя (например, " + Chr(34) + fName + "1" + Chr(34) + ")!"
                Return -1
            End If
        End If

        'получаем Id книги заклинаний а нобора способностей героя
        Dim mBook As String = "", abSet As String = "", mBookId As Integer = -1, abSetId As Integer = -1
        If ReadProperty(classH, "MagicBook", heroId, -1, mBook, arrHeroId, MatewScript.ReturnFormatEnum.ORIGINAL, -1, True) = False Then Return -1
        If ReadProperty(classH, "AbilitiesSet", heroId, -1, abSet, arrHeroId, MatewScript.ReturnFormatEnum.ORIGINAL, -1, True) = False Then Return -1
        mBookId = Math.Max(GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties), -1)
        abSetId = Math.Max(GetSecondChildIdByName(abSet, mScript.mainClass(classAb).ChildProperties), -1)
        'создаем бойца
        Dim fig As New cFighters With {.heroId = heroId, .magicBookId = mBookId, .abilitySetId = abSetId}

        'не является ли персонаж дубликатом (второй слон, десятая крыса ...)
        Dim srcProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classH).ChildProperties(heroId)
        Dim oldId As Integer = GetFighterByName(srcProps("Name").Value), isDuplicate As Boolean = False
        If oldId > -1 Then isDuplicate = True 'да, является

        'копируем свойства от героя к бойцу
        fig.heroProps = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)
        Dim i As Integer
        For i = 0 To srcProps.Count - 1
            fig.heroProps.Add(srcProps.ElementAt(i).Key, srcProps.ElementAt(i).Value.Clone)
        Next
        fig.heroProps("Name").Value = fName
        fig.heroProps("Army").Value = armyId.ToString
        fig.heroProps("HeroOwner").Value = owner.ToString
        If owner > -1 Then fig.heroProps("Friendliness").Value = "2"

        If isDuplicate Then
            fig.isDuplicate = True
            'создаем дубликаты книги заклинаний и набора способностей
            If fig.magicBookId > -1 Then
                i = 1
                Do
                    mBook = "'MagicBook" & i.ToString & "'"
                    If ObjectIsExists(classMg, {mBook}) <> "True" Then Exit Do
                Loop
                mBook = CreateNewObject(classMg, {mBook})
                If mBook = "#Error" Then Return -1
                mBookId = Val(mBook)

                srcProps = mScript.mainClass(classMg).ChildProperties(fig.magicBookId)
                Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classMg).ChildProperties(mBookId)
                For i = 0 To srcProps.Count - 1
                    'Dim ch As MatewScript.ChildPropertiesInfoType = destProps.ElementAt(i).Value
                    Dim key As String = srcProps.ElementAt(i).Key
                    Dim ch As MatewScript.ChildPropertiesInfoType = srcProps.ElementAt(i).Value.Clone
                    destProps(key) = ch
                Next i

                fig.heroProps("MagicBook").Value = mBookId.ToString
                fig.magicBookId = mBookId
            End If

            If fig.abilitySetId > -1 Then
                i = 1
                Do
                    abSet = "'AbilitiesSet" & i.ToString & "'"
                    If ObjectIsExists(classAb, {abSet}) <> "True" Then Exit Do
                    i += 1
                Loop
                abSet = CreateNewObject(classAb, {abSet})
                If abSet = "#Error" Then Return -1
                abSetId = Val(abSet)

                srcProps = mScript.mainClass(classAb).ChildProperties(fig.abilitySetId)
                Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classAb).ChildProperties(abSetId)
                For i = 0 To srcProps.Count - 1
                    'Dim ch As MatewScript.ChildPropertiesInfoType = destProps.ElementAt(i).Value
                    Dim key As String = srcProps.ElementAt(i).Key
                    Dim ch As MatewScript.ChildPropertiesInfoType = srcProps.ElementAt(i).Value.Clone
                    destProps(key) = ch
                Next i

                fig.heroProps("AbilitiesSet").Value = abSetId.ToString
                fig.abilitySetId = abSetId
            End If
        Else
            'герой - не дубликат. Для возможности восстановления сохраняем исходное состояние его книги заклианий и набора способностей
            If fig.magicBookId > -1 Then
                'сохраняем книгу магий
                srcProps = mScript.mainClass(classMg).ChildProperties(fig.magicBookId)
                Dim destProps As New SortedList(Of String, MatewScript.ChildPropertiesInfoType)
                For i = 0 To srcProps.Count - 1
                    Dim key As String = srcProps.ElementAt(i).Key
                    Dim ch As MatewScript.ChildPropertiesInfoType = srcProps.ElementAt(i).Value.Clone
                    destProps.Add(key, ch)
                Next i
                lstMagicBooksCopies.Add(fig.heroId, destProps)
            End If

            If fig.abilitySetId > -1 Then
                'сохраняем набор способностей
                srcProps = mScript.mainClass(classAb).ChildProperties(fig.abilitySetId)
                Dim destProps As New SortedList(Of String, MatewScript.ChildPropertiesInfoType)
                For i = 0 To srcProps.Count - 1
                    Dim key As String = srcProps.ElementAt(i).Key
                    Dim ch As MatewScript.ChildPropertiesInfoType = srcProps.ElementAt(i).Value.Clone
                    destProps.Add(key, ch)
                Next i
                lstAbilityCopies.Add(fig.heroId, destProps)
            End If
        End If

        'If G_VAR.G_ISBATTLE = True Then ReDim Preserve arrBattleTurns(UBound(arrBattleTurns) + 1)
        fig.armyId = armyId
        fig.armyUnitId = armyUnitId
        Fighters.Add(fig)

        If GVARS.G_ISBATTLE = False AndAlso armyId > -1 Then
            Dim classArmy As Integer = mScript.mainClassHash("Army")
            'изменяем начальное количество бойцов армии
            Dim initCount As Integer = 0
            If ReadPropertyInt(classArmy, "InitUnitsCount", armyId, -1, initCount, {armyId.ToString}) = False Then Return -1
            initCount += 1
            If PropertiesRouter(classArmy, "InitUnitsCount", {armyId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, initCount.ToString) = "#Error" Then Return -1
        End If

        Return (Fighters.Count - 1)
    End Function

    ''' <summary>
    ''' Заявляет на бой всех бойцов данной армии
    ''' </summary>
    ''' <param name="armyId">Id армии</param>
    ''' <returns>Количество добавленных юнитов</returns>
    Public Function AddArmy(ByVal armyId As Integer) As Integer
        Dim classArmy As Integer = mScript.mainClassHash("Army"), classH As Integer = mScript.mainClassHash("H")
        If IsNothing(mScript.mainClass(classArmy).ChildProperties(armyId)("Name").ThirdLevelProperties) OrElse mScript.mainClass(classArmy).ChildProperties(armyId)("Name").ThirdLevelProperties.Count = 0 Then Return ""

        Dim hero As String = "", res As String = ""
        Dim arrs() As String = {armyId.ToString, ""}
        Dim globalEventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnBattleApplication").eventId
        Dim armyEventId As Integer = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnBattleApplication").eventId
        Dim unitsAdded As Integer = 0
        For uId As Integer = 0 To mScript.mainClass(classArmy).ChildProperties(armyId)("Name").ThirdLevelProperties.Count - 1
            arrs(1) = uId.ToString
            If ReadProperty(classArmy, "Hero", armyId, uId, hero, arrs) = False Then Return -1
            Dim heroId As Integer = GetSecondChildIdByName(hero, mScript.mainClass(classH).ChildProperties)
            If heroId < 0 Then
                mScript.LAST_ERROR = "Персонажа " & hero & " не существует."
                Return -1
            End If
            Dim unitCount As Integer = 1, participate As Boolean = True
            If ReadPropertyBool(classArmy, "ParticipateInBattle", armyId, uId, participate, arrs) = False Then Return -1
            If participate = False Then Continue For
            If ReadPropertyInt(classArmy, "CountInBattle", armyId, uId, unitCount, arrs) = False Then Return -1

            'Событие ArmyOnBattleApplication
            If globalEventId > 0 Then
                'глобальное
                res = mScript.eventRouter.RunEvent(globalEventId, arrs, "ArmyOnBattleApplication", False)
                If res = "False" Then
                    Continue For
                ElseIf IsNumeric(res) Then
                    unitCount = Val(res)
                End If
            End If

            If armyEventId > 0 Then
                'данной армии
                res = mScript.eventRouter.RunEvent(armyEventId, arrs, "ArmyOnBattleApplication", False)
                If res = "False" Then
                    Continue For
                ElseIf IsNumeric(res) Then
                    unitCount = Val(res)
                End If
            End If

            Dim eventId As Integer = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnBattleApplication").ThirdLevelEventId(uId)
            If eventId > 0 Then
                'данного юнита
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnBattleApplication", False)
                If res = "False" Then
                    Continue For
                ElseIf IsNumeric(res) Then
                    unitCount = Val(res)
                End If
            End If

            If unitCount <= 0 Then Continue For
            Dim capt As String = "", range As Integer = 0
            If ReadProperty(classArmy, "Caption", armyId, uId, capt, arrs) = False Then Return -1
            If ReadPropertyInt(classArmy, "Range", armyId, uId, range, arrs) = False Then Return -1

            For j As Integer = 1 To unitCount
                Dim fId As Integer = AddFighter(heroId, "", armyId.ToString, uId)
                If fId < 0 Then Return -1
                Dim props As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Fighters(fId).heroProps
                If String.IsNullOrEmpty(capt) OrElse capt <> "''" Then props("Caption").Value = capt
                props("Range").Value = range.ToString
                unitsAdded += 1
            Next j
        Next uId

        If GVARS.G_ISBATTLE Then
            'изменяем начальное количество бойцов армии
            Dim initCount As Integer = 0
            If ReadPropertyInt(classArmy, "InitUnitsCount", armyId, -1, initCount, {armyId.ToString}) = False Then Return -1
            initCount += unitsAdded
            If PropertiesRouter(classArmy, "InitUnitsCount", {armyId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, initCount.ToString) = "#Error" Then Return -1
        End If

        Return unitsAdded
    End Function

    ''' <summary>
    ''' Заявляет на бой бойца армии
    ''' </summary>
    ''' <param name="armyId">Id армии</param>
    ''' <param name="unitId">Id бойца</param>
    ''' <returns>-1 при ошибке или количество добавленных бойцов</returns>
    Public Function AddArmyUnit(ByVal armyId As Integer, ByVal unitId As Integer, Optional ByVal owner As Integer = -1) As Integer
        Dim classArmy As Integer = mScript.mainClassHash("Army"), classH As Integer = mScript.mainClassHash("H")

        Dim hero As String = "", res As String = ""
        Dim arrs() As String = {armyId.ToString, unitId.ToString}
        Dim globalEventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnBattleApplication").eventId
        Dim armyEventId As Integer = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnBattleApplication").eventId
        If ReadProperty(classArmy, "Hero", armyId, unitId, hero, arrs) = False Then Return -1
        Dim heroId As Integer = GetSecondChildIdByName(hero, mScript.mainClass(classH).ChildProperties)
        If heroId < 0 Then
            _ERROR("Персонажа " & hero & " не существует.")
            Return -1
        End If
        Dim unitCount As Integer = 1, participate As Boolean = True
        If ReadPropertyBool(classArmy, "ParticipateInBattle", armyId, unitId, participate, arrs) = False Then Return -1
        If participate = False Then Return 0
        If ReadPropertyInt(classArmy, "CountInBattle", armyId, unitId, unitCount, arrs) = False Then Return -1

        'Событие ArmyOnBattleApplication
        If globalEventId > 0 Then
            'глобальное
            res = mScript.eventRouter.RunEvent(globalEventId, arrs, "ArmyOnBattleApplication", False)
            If res = "False" Then
                Return 0
            ElseIf IsNumeric(res) Then
                unitCount = Val(res)
            End If
        End If

        If armyEventId > 0 Then
            'данной армии
            res = mScript.eventRouter.RunEvent(armyEventId, arrs, "ArmyOnBattleApplication", False)
            If res = "False" Then
                Return 0
            ElseIf IsNumeric(res) Then
                unitCount = Val(res)
            End If
        End If

        Dim eventId As Integer = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnBattleApplication").ThirdLevelEventId(unitId)
        If eventId > 0 Then
            'данного юнита
            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnBattleApplication", False)
            If res = "False" Then
                Return 0
            ElseIf IsNumeric(res) Then
                unitCount = Val(res)
            End If
        End If

        If unitCount <= 0 Then Return 0
        Dim capt As String = "", range As String = ""
        If ReadProperty(classArmy, "Caption", armyId, unitId, capt, arrs) = False Then Return -1
        If ReadProperty(classArmy, "Range", armyId, unitId, range, arrs) = False Then Return -1

        Dim fId As Integer = -1
        For j As Integer = 1 To unitCount
            fId = AddFighter(heroId, "", armyId.ToString, unitId, owner)
            If fId < 0 Then Return -1
            Dim props As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Fighters(fId).heroProps
            If String.IsNullOrEmpty(capt) OrElse capt <> "''" Then props("Caption").Value = capt
            props("Range").Value = range.ToString
            If GVARS.G_ISBATTLE Then
                If AppendFighterInBattle(fId) = "#Error" Then Return "#Error"
            End If
        Next j

        Return unitCount
    End Function

    'Public Function RemoveArmy

    ''' <summary>
    ''' Удаляет бойца, заявленного на бой
    ''' </summary>
    ''' <param name="fId">Id </param>
    ''' <returns>Id бойца</returns>
    Public Function RemoveFighter(ByVal fId As Integer) As String
        If GVARS.G_ISBATTLE Then
            Return _ERROR("Нельзя удалить бойца во время битвы! Он может только умереть или сбежать (Life = 0 или RunAway = True)!")
        End If

        Dim classH As Integer = mScript.mainClassHash("H")
        If fId < 0 Then
            'удаляем всех бойцов
            If Fighters.Count = 0 Then Return ""
            For i As Integer = Fighters.Count - 1 To 0 Step -1
                If Fighters(i).isDuplicate Then
                    'удаляем дубликаты книги заклинаний и набора способностей
                    If Fighters(i).magicBookId > -1 Then
                        Dim classMg As Integer = mScript.mainClassHash("Mg"), mBookId As Integer = Fighters(i).magicBookId
                        If RemoveObject(classMg, {mBookId.ToString}) = "#Error" Then Return "#Error"
                        'уменьшаем на 1 Id книг заклинаний у других героев, если этот Id был больше удаленного mBookId
                        For j As Integer = 0 To Fighters.Count - 1
                            Dim curMBook As Integer = Fighters(j).magicBookId
                            If curMBook > mBookId Then
                                If PropertiesRouter(classH, "MagicBook", {j.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, (curMBook - 1).ToString) = "#Error" Then Return "#Error"
                            End If
                        Next j
                    End If
                    If Fighters(i).abilitySetId > -1 Then
                        Dim classAb As Integer = mScript.mainClassHash("Ab"), abSetId As Integer = Fighters(i).abilitySetId
                        If RemoveObject(classAb, {Fighters(i).abilitySetId.ToString}) = "#Error" Then Return "#Error"
                        'уменьшаем на 1 Id наборов способностей у других героев, если этот Id был больше удаленного abSetId
                        For j As Integer = 0 To Fighters.Count - 1
                            Dim curAbSet As Integer = Fighters(j).abilitySetId
                            If curAbSet > abSetId Then
                                If PropertiesRouter(classH, "AbilitiesSet", {j.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, (curAbSet - 1).ToString) = "#Error" Then Return "#Error"
                            End If
                        Next j
                    End If
                End If
                'удаляем бойца
                Fighters.RemoveAt(i)
            Next i
        Else
            'удаляем бойца fId
            If Fighters(fId).isDuplicate Then
                'удаляем дубликаты книги заклинаний и набора способностей
                If Fighters(fId).magicBookId > -1 Then
                    Dim classMg As Integer = mScript.mainClassHash("Mg")
                    If RemoveObject(classMg, {Fighters(fId).magicBookId.ToString}) = "#Error" Then Return "#Error"
                    'уменьшаем на 1 Id книг заклинаний у других героев, если этот Id был больше удаленного mBookId
                    For j As Integer = 0 To Fighters.Count - 1
                        Dim curMBook As Integer = Fighters(j).magicBookId
                        If curMBook > Fighters(fId).magicBookId Then
                            If PropertiesRouter(classH, "MagicBook", {j.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, (curMBook - 1).ToString) = "#Error" Then Return "#Error"
                        End If
                    Next j
                End If
                If Fighters(fId).abilitySetId > -1 Then
                    Dim classAb As Integer = mScript.mainClassHash("Ab")
                    If RemoveObject(classAb, {Fighters(fId).abilitySetId.ToString}) = "#Error" Then Return "#Error"
                    'уменьшаем на 1 Id наборов способностей у других героев, если этот Id был больше удаленного abSetId
                    For j As Integer = 0 To Fighters.Count - 1
                        Dim curAbSet As Integer = Fighters(j).abilitySetId
                        If curAbSet > Fighters(fId).abilitySetId Then
                            If PropertiesRouter(classH, "AbilitiesSet", {j.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, (curAbSet - 1).ToString) = "#Error" Then Return "#Error"
                        End If
                    Next j
                End If
            End If
            'удаляем бойца
            Fighters.RemoveAt(fId)
        End If
        Return ""
    End Function

#End Region

    ''' <summary>Начинает бой</summary>
    Public Function Begin() As String
        If GVARS.G_ISBATTLE Then Return _ERROR("Battle.Start не возможно! Битва уже идет!")
        If Fighters.Count = 0 Then
            Return _ERROR("Battle.Start не возможно! На битву не зарезервирован ни один боец!")
        ElseIf Fighters.Count = 1 Then
            Return _ERROR("Battle.Start не возможно! Один боец не может драться сам с собой!")
        End If

        lstHistory.Clear()
        'Удаляем все события
        If mScript.ExecuteString({"A.Remove"}, Nothing) = "#Error" Then Return "#Error"
        'Turnscount = 0
        Dim classB As Integer = mScript.mainClassHash("B")
        If PropertiesRouter(classB, "TurnsCount", Nothing, Nothing, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return "#Error"

        GVARS.G_ISBATTLE = True
        GVARS.G_CURVICTIM = -1
        GVARS.G_STRIKETYPE = -1
        'G_VAR.G_BATTLEDURATION = -1
        GVARS.G_CANUSEMAGIC = False
        GVARS.TIME_IN_BATTLE = Now.Ticks

        'Вывод поля боя
        If PrepareBattlefield() = "#Error" Then Return "#Error"

        'get starting hero
        Dim startingHero As String = "", startingHeroId As Integer = -1
        If ReadProperty(classB, "StartingHero", -1, -1, startingHero, Nothing) = False Then Return "#Error"
        startingHeroId = GetFighterByName(startingHero)
        If startingHeroId < 0 Then
            GVARS.G_CURFIGHTER = GetNextFighter(-1, False)
            If GVARS.G_CURFIGHTER < 0 Then Return _ERROR("Battle.Start не возможно! Не найдено ни одного живого бойца!", "Battle Begin")
        Else
            GVARS.G_CURFIGHTER = startingHeroId
        End If

        'добавляем css действий в главное окно (для корректного вывода действий)
        Dim actCSS As String = ""
        ReadProperty(mScript.mainClassHash("AW"), "CSS", -1, -1, actCSS, Nothing)
        actCSS = UnWrapString(actCSS)
        If String.IsNullOrEmpty(actCSS) = False Then HtmlAppendCSS(frmPlayer.wbMain.Document, actCSS)

        'Запускаем событие BattleOnShowPrepare
        Dim eventId As Integer = mScript.mainClass(classB).Properties("BattleOnShowPrepare").eventId
        If eventId > 0 Then
            If mScript.eventRouter.RunEvent(eventId, Nothing, "BattleOnShowPrepare", False) = "#Error" Then Return "#Error"
        End If


        'ReDim arrBattleTurns(UBound(masFighter))
        'battle_Finished = False
        Return ""
    End Function

    ''' <summary>Завершает бой</summary>
    Public Function StopBattle(Optional ByVal reason As BattleFinishEnum = BattleFinishEnum.CHECK_ERROR) As String
        If reason = BattleFinishEnum.CHECK_ERROR Then reason = CheckForBattleFinished()
        If reason = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
        Dim battleResult As String = Val(reason).ToString

        If mScript.ExecuteString({"A.Remove"}, Nothing) = "#Error" Then Return "#Error"
        'Событие BattleOnFinishEvent
        Dim classB As Integer = mScript.mainClassHash("B"), res As String = ""
        Dim eventId As Integer = mScript.mainClass(classB).Properties("BattleOnFinishEvent").eventId
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, {battleResult}, "BattleOnFinishEvent", False)
            If res = "False" Then
                Return res
            ElseIf res = "#Error" Then
                Return res
            End If
        End If

        'проверка списка воспроизведения
        If GVARS.G_CURLIST > -1 AndAlso GVARS.G_CURLIST <> audioListBeforeBattle Then
            If AudioStop() = False Then Return "#Error"
            GVARS.G_CURAUDIO = -1
            GVARS.G_CURLIST = audioListBeforeBattle
            If GVARS.G_CURLIST > -1 Then
                If AudioPlayFromList(-1, Nothing) = "#Error" Then Return "#Error"
                timAudio.Enabled = True
            End If
        End If

        'Восстанавливаем переменные боя 
        GVARS.TIME_IN_BATTLE = 0
        GVARS.G_CURFIGHTER = -1
        GVARS.G_CURVICTIM = -1
        GVARS.G_STRIKETYPE = -1
        GVARS.G_ISBATTLE = False

        'Восстанавливаем свойство Battle.TurnsCount 
        If PropertiesRouter(classB, "TurnsCount", Nothing, Nothing, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return "#Error"

        'Восстанавливаем свойства героев с SaveBattleResults = True и возвращаем исходные книги заклинаний и наборы способностей при SaveBattleResults = False
        Dim classMg As Integer = mScript.mainClassHash("Mg"), classAb As Integer = mScript.mainClassHash("Ab"), classH As Integer = mScript.mainClassHash("H")
        For fId As Integer = 0 To Fighters.Count - 1
            If Fighters(fId).isDuplicate Then
                'боец-дубликат. Удаляем его книгу заклинаний и набор способностей
                If Fighters(fId).magicBookId > -1 Then
                    Dim mbId As Integer = Fighters(fId).magicBookId
                    If RemoveObject(classMg, {mbId.ToString}) = "#Error" Then Return "#Error"
                    'уменьшаем magicBookId других книг на 1 если magicBookId > mbId
                    For j As Integer = 0 To Fighters.Count - 1
                        If Fighters(j).magicBookId > mbId Then Fighters(j).magicBookId -= 1
                    Next j
                End If

                If Fighters(fId).abilitySetId > -1 Then
                    Dim abSetId As Integer = Fighters(fId).abilitySetId
                    If RemoveObject(classAb, {abSetId.ToString}) = "#Error" Then Return "#Error"
                    'уменьшаем abilitySetId других наборов на 1 если abilitySetId > abSetId
                    For j As Integer = 0 To Fighters.Count - 1
                        If Fighters(j).abilitySetId > abSetId Then Fighters(j).abilitySetId -= 1
                    Next j
                End If

                'очищаем события всех свойств данного бойца
                For pId As Integer = 0 To Fighters(fId).heroProps.Count - 1
                    eventId = Fighters(fId).heroProps.ElementAt(pId).Value.eventId
                    If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId)
                Next pId
                Continue For
            End If

            Dim saveResults As Boolean = False
            If ReadFighterPropertyBool("SaveBattleResults", fId, saveResults, {fId.ToString}) = False Then Return "#Error"
            If saveResults Then
                'Результаты боя сохраняются. Надо установить свойства из Fighters в Hero, сбросить Mg.SpelledThisBattle и Ab.TurnsActive и удалить способности с RemoveAfterBattle = True
                If Fighters(fId).abilitySetId > -1 Then
                    Dim abSetId As Integer = Fighters(fId).abilitySetId
                    'Удаляем способности с RemoveAfterBattle (с выполнением AbilityOnReleaseEvent)
                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityOnReleaseEvent")
                    If IsNothing(ch.ThirdLevelProperties) = False Then
                        For abId As Integer = ch.ThirdLevelProperties.Count - 1 To 0 Step -1
                            Dim shouldRemove As Boolean = False
                            If ReadPropertyBool(classAb, "RemoveAfterBattle", abSetId, abId, shouldRemove, {abSetId.ToString, abId.ToString}) = False Then Return "#Error"
                            If Not shouldRemove Then Continue For
                            If FunctionRouter(classAb, "Remove", {abSetId.ToString, abId.ToString}, Nothing) = "#Error" Then Return "#Error"
                        Next abId
                    End If

                    'Устанавливаем свойства из Fighters в Hero, сбрасываем Mg.SpelledThisBattle и Ab.TurnsActive
                    If SaveFighterToHero(fId) = "#Error" Then Return "#Error"
                End If
            Else
                'Результаты боя не сохраняются. Надо восстановить исходные книгу заклинаний и набор способностей. Удалять способности с вызовом AbilityOnReleaseEvent не нужно
                Dim heroId As Integer = Fighters(fId).heroId
                Dim arrs() As String = {heroId.ToString}
                If lstMagicBooksCopies.ContainsKey(heroId) Then
                    'книга магий была. Получаем ее исходный Id (на случай, если персонажу выдали новую книгу во время боя)
                    Dim initMBook As String = "", initMBookId As Integer = -1
                    If ReadProperty(classH, "MagicBook", heroId, -1, initMBook, arrs, MatewScript.ReturnFormatEnum.ORIGINAL, -1, True) = False Then Return "#Error"
                    initMBookId = GetSecondChildIdByName(initMBook, mScript.mainClass(classMg).ChildProperties)
                    If initMBookId < 0 Then Return _ERROR("Сбой при восстановлении книги магий героя с Id " & heroId.ToString & ". Книга заклинаний " & initMBook & " не найдена.")

                    Dim srcProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstMagicBooksCopies(heroId)
                    Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classMg).ChildProperties(initMBookId)
                    For pId As Integer = 0 To srcProps.Count - 1
                        Dim key As String = srcProps.ElementAt(pId).Key
                        'очищаем старые события
                        Dim chDest As MatewScript.ChildPropertiesInfoType = destProps(key)
                        If chDest.eventId > 0 Then mScript.eventRouter.RemoveEvent(chDest.eventId)
                        If IsNothing(chDest.ThirdLevelEventId) = False AndAlso chDest.ThirdLevelEventId.Count > 0 Then
                            For child3Id As Integer = 0 To chDest.ThirdLevelEventId.Count - 1
                                If chDest.ThirdLevelEventId(child3Id) > 0 Then mScript.eventRouter.RemoveEvent(chDest.ThirdLevelEventId(child3Id))
                            Next child3Id
                        End If
                        'устанавливаем новое свойство
                        destProps(key) = srcProps(key)
                    Next pId
                End If

                If lstAbilityCopies.ContainsKey(heroId) Then
                    'набор способостей был. Получаем его исходный Id (на случай, если персонажу выдали новый набор во время боя)
                    Dim initAbSet As String = "", initAbSetId As Integer = -1
                    If ReadProperty(classH, "AbilitiesSet", heroId, -1, initAbSet, arrs, MatewScript.ReturnFormatEnum.ORIGINAL, -1, True) = False Then Return "#Error"
                    initAbSetId = GetSecondChildIdByName(initAbSet, mScript.mainClass(classAb).ChildProperties)
                    If initAbSetId < 0 Then Return _ERROR("Сбой при восстановлении набора способостей героя с Id " & heroId.ToString & ". Набор способостей " & initAbSetId & " не найден.")

                    Dim srcProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstAbilityCopies(heroId)
                    Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classAb).ChildProperties(initAbSetId)
                    For pId As Integer = 0 To srcProps.Count - 1
                        Dim key As String = srcProps.ElementAt(pId).Key
                        'очищаем старые события
                        Dim chDest As MatewScript.ChildPropertiesInfoType = destProps(key)
                        If chDest.eventId > 0 Then mScript.eventRouter.RemoveEvent(chDest.eventId)
                        If IsNothing(chDest.ThirdLevelEventId) = False AndAlso chDest.ThirdLevelEventId.Count > 0 Then
                            For child3Id As Integer = 0 To chDest.ThirdLevelEventId.Count - 1
                                If chDest.ThirdLevelEventId(child3Id) > 0 Then mScript.eventRouter.RemoveEvent(chDest.ThirdLevelEventId(child3Id))
                            Next child3Id
                        End If
                        'устанавливаем новое свойство
                        destProps(key) = srcProps(key)
                    Next pId
                End If

                'очищаем события всех свойств данного бойца
                For pId As Integer = 0 To Fighters(fId).heroProps.Count - 1
                    eventId = Fighters(fId).heroProps.ElementAt(pId).Value.eventId
                    If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId)
                Next pId
            End If
        Next fId

        'очищаем данные о бое
        lstAbilityCopies.Clear()
        lstMagicBooksCopies.Clear()
        lstPositions.Clear()
        lstVictims.Clear()
        audioListBeforeBattle = -1
        Fighters.Clear()
        JustAfterBattle = True

        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть html-документ Battle.html!")
        Dim hConvas As HtmlElement = hDoc.GetElementById("BattleConvas")
        If IsNothing(hConvas) Then Return _ERROR("Нарушена структура html-документа Battle.html. Элемент с Id BattleConvas не найден!")
        hConvas.InnerHtml = ""
        Dim hHistory As HtmlElement = hDoc.GetElementById("MainConvas")
        If IsNothing(hHistory) = False Then
            hHistory.Style &= "left:0px;top:0px;width:calc(100% - 20px);"
        End If
        ''change style
        'bStr = bBackColor
        'If Len(bStr) = 0 Then bStr = "Empty"
        'aStr = "LW.BackColor = " + bStr + vbNewLine
        'bStr = bBackPicture
        'If Len(bStr) = 0 Then bStr = "Empty"
        'aStr = aStr + "LW.BackPicture = " + bStr + vbNewLine
        'bStr = bBackPicStyle
        'If Len(bStr) = 0 Then bStr = "Empty"
        'aStr = aStr + "LW.BackPicStyle = " + bStr + vbNewLine
        'bStr = bBackPicPos
        'If Len(bStr) = 0 Then bStr = "Empty"
        'aStr = aStr + "LW.BackPicPos = " + bStr
        'DoBlockInt(aStr, 0)

        ''show exit
        'aStr = "LW.Clear" + vbNewLine + "A.Remove"
        'DoBlockInt(aStr, 0)

        'Событие BattleOnExitEvent
        eventId = mScript.mainClass(classB).Properties("BattleOnExitEvent").eventId
        If eventId > -1 Then
            res = mScript.eventRouter.RunEvent(eventId, {battleResult}, "BattleOnExitEvent", False)
            If res = "#Error" Then Return res
        End If
        killingSummoned = False

        'сбрасываем начальное количество бойцов армии        
        Dim classArmy As Integer = mScript.mainClassHash("Army")
        If IsNothing(mScript.mainClass(classArmy).ChildProperties) = False AndAlso mScript.mainClass(classArmy).ChildProperties.Count > 0 Then
            For i As Integer = 0 To mScript.mainClass(classArmy).ChildProperties.Count - 1
                If PropertiesRouter(classArmy, "InitUnitsCount", {i.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return -1
            Next i
        End If

        Return ""
    End Function

    ''' <summary>Завершает бой при закрытии квеста, когда не надо запускать никакие события и целостность данных структуры квеста не имеет значения.</summary>
    Public Sub StopBattleAbrupt()
        'Восстанавливаем переменные боя 
        GVARS.TIME_IN_BATTLE = 0
        GVARS.G_CURFIGHTER = -1
        GVARS.G_CURVICTIM = -1
        GVARS.G_STRIKETYPE = -1
        GVARS.G_ISBATTLE = False

        'очищаем данные о бое
        lstAbilityCopies.Clear()
        lstMagicBooksCopies.Clear()
        lstPositions.Clear()
        lstVictims.Clear()
        audioListBeforeBattle = -1
        Fighters.Clear()
        killingSummoned = False
    End Sub

    ''' <summary>Проверяет не наступило ли завершение битвы (по количеству живых бойцов)</summary>
    Private Function CheckForBattleFinished() As BattleFinishEnum
        Dim enemies As Integer = 0, mainHeroDead As Boolean = False
        Dim fr As Integer = 0, life As Double = 0, runAway As Boolean = False
        For fId As Integer = 0 To Fighters.Count - 1
            Dim arr() As String = {fId.ToString}
            If ReadFighterPropertyInt("Friendliness", fId, fr, arr) = False Then Return BattleFinishEnum.CHECK_ERROR
            'If fr = 2 Then
            '    'если боец призванный, то получаем дружественность хозяина
            '    Dim fHost As Integer = GetSummonOwner(fId)
            '    If ReadFighterPropertyInt("Friendliness", fHost, fr, {fHost.ToString}) = False Then Return BattleFinishEnum.CHECK_ERROR
            'End If

            If ReadFighterPropertyDbl("Life", fId, life, arr) = False Then Return BattleFinishEnum.CHECK_ERROR
            If ReadFighterPropertyBool("RunAway", fId, runAway, arr) = False Then Return BattleFinishEnum.CHECK_ERROR
            If (life <= 0 OrElse runAway) Then
                If fr = -1 Then mainHeroDead = True
                Continue For
            End If
            If IsEnemy(-1, fId) = IsEnemyEnum.Enemy Then enemies += 1
        Next fId

        If enemies = 0 AndAlso mainHeroDead Then
            Return BattleFinishEnum.ALL_IS_DEAD
        ElseIf enemies = 0 Then
            Return BattleFinishEnum.ENEMIES_DEAD
        ElseIf mainHeroDead Then
            Return BattleFinishEnum.MAIN_HERO_DEAD
        Else
            Return BattleFinishEnum.CUSTOM
        End If
    End Function

    ''' <summary>
    ''' Устанавливает все свойства бойца на соответствующего ему герою (для SaveBattleResults = True) и сбрасывает счетчики Mg.SpelledThisBattle и Ab.TurnsActive магий и способностей героя
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    Public Function SaveFighterToHero(ByVal fId As Integer) As String
        If Fighters(fId).isDuplicate Then Return _ERROR("Нельзя восстанавливать копию героя из дубликата!")

        Dim heroId As Integer = Fighters(fId).heroId
        Dim classH As Integer = mScript.mainClassHash("H"), arrs() As String = {heroId.ToString}
        Dim srcProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Fighters(fId).heroProps
        Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classH).ChildProperties(heroId)
        'свойства, которые не надо устанавливать:
        Dim propsToExclude As List(Of String) = {"Name", "RunAway", "Army", "Range", "HeroOwner", "AttacksInThisTurn", "AbilitiesSet", "MagicBook", "Friendliness"}.ToList
        For pId As Integer = 0 To srcProps.Count - 1
            Dim key As String = srcProps.ElementAt(pId).Key 'имя свойства
            If propsToExclude.IndexOf(key) > -1 Then Continue For
            If destProps(key).Value <> srcProps(key).Value Then
                If PropertiesRouter(classH, key, arrs, Nothing, PropertiesOperationEnum.PROPERTY_SET, srcProps(key).Value, True) = "#Error" Then Return "#Error"
            End If
        Next pId

        Dim classMg As Integer = mScript.mainClassHash("Mg"), classAb As Integer = mScript.mainClassHash("Ab")
        If Fighters(fId).magicBookId > -1 Then
            'восстанавливаем исходное значение Mg.SpelledThisBattle
            Dim initValue As String = mScript.mainClass(classMg).Properties("SpelledThisBattle").Value
            Dim mBookId As Integer = Fighters(fId).magicBookId
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMg).ChildProperties(mBookId)("SpelledThisBattle")
            If IsNothing(ch.ThirdLevelProperties) = False AndAlso ch.ThirdLevelProperties.Count > 0 Then
                For i As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                    If PropertiesRouter(classMg, "SpelledThisBattle", {mBookId.ToString, i.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, initValue) = "#Error" Then Return "#Error"
                Next i
            End If
        End If

        If Fighters(fId).abilitySetId > -1 Then
            'восстанавливаем исходное значение Ab.TurnsActive
            Dim initValue As String = mScript.mainClass(classAb).Properties("TurnsActive").Value
            Dim abSetId As Integer = Fighters(fId).abilitySetId
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("TurnsActive")
            If IsNothing(ch.ThirdLevelProperties) = False AndAlso ch.ThirdLevelProperties.Count > 0 Then
                For i As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                    If PropertiesRouter(classAb, "TurnsActive", {abSetId.ToString, i.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, initValue) = "#Error" Then Return "#Error"
                Next i
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Выводит на экран поле боя со всеми армиями
    ''' </summary>
    Private Function PrepareBattlefield() As String
        'Загружаем Battle.html
        Dim f As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Battle.html")
        If FileIO.FileSystem.FileExists(f) = False Then
            Return _ERROR("В папке квеста не найден файл " & f & ", необходимый для запуска игры. Попробуйте скопировать его вручную из " & _
                            FileIO.FileSystem.CombinePath(Application.StartupPath, "src\defaultFiles") & " .")
        End If
        frmPlayer.wbMain.AllowNavigation = True
        frmPlayer.wbMain.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Battle.html"))
        Do Until frmPlayer.wbMain.ReadyState = WebBrowserReadyState.Complete
            Application.DoEvents()
        Loop

        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть html-документ Battle.html!")
        Dim hConvas As HtmlElement = hDoc.GetElementById("BattleConvas")
        If IsNothing(hConvas) Then Return _ERROR("Нарушена структура html-документа Battle.html. Элемент с Id BattleConvas не найден!")
        frmPlayer.wbMain.AllowNavigation = False
        Dim classB As Integer = mScript.mainClassHash("B"), classArmy As Integer = mScript.mainClassHash("Army")

        'устанавливаем фон
        Dim strStyle As New System.Text.StringBuilder
        Dim bgColor As String = ""
        ReadProperty(classB, "BackColor", -1, -1, bgColor, Nothing)
        bgColor = UnWrapString(bgColor)
        If String.IsNullOrEmpty(bgColor) = False Then strStyle.Append("background:" & bgColor & ";")

        'Меняем фоновую картинку
        Dim bkPicture As String = ""
        ReadProperty(classB, "BackPicture", -1, -1, bkPicture, Nothing)
        bkPicture = UnWrapString(bkPicture)
        Dim bkPicPos As Integer = 0, bkPicStyle As Integer = 0
        If String.IsNullOrEmpty(bkPicture) = False Then
            bkPicture = bkPicture.Replace("\"c, "/"c)
            ReadPropertyInt(classB, "BackPicPos", -1, -1, bkPicPos, Nothing)
            ReadPropertyInt(classB, "BackPicStyle", -1, -1, bkPicStyle, Nothing)
        End If

        If String.IsNullOrEmpty(bkPicture) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)) Then
            'файл картинки существует
            strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture + "'")
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Select Case bkPicStyle
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select
            If bkPicStyle = 0 Then
                'BackPicPos
                '0 в левом верхнем углу, 1 слева по центру, 2 в левом нижнем углу, 3 сверху по центру, 4 в центре, 5 снизу по центру, 6 в правом верхнем углу, 7 справа по центру, 8 в правом нижнем углу
                strStyle.Append("background-position:")
                Select Case bkPicPos
                    Case 0 '0 в левом верхнем углу
                        strStyle.Append("left top;")
                    Case 1 '1 слева по центру
                        strStyle.Append("left center;")
                    Case 2 '2 в левом нижнем углу
                        strStyle.Append("left bottom;")
                    Case 3 '3 сверху по центру
                        strStyle.Append("center top;")
                    Case 4 '4 в центре
                        strStyle.Append("center center;")
                    Case 5 '5 снизу по центру
                        strStyle.Append("center bottom;")
                    Case 6 '6 в правом верхнем углу
                        strStyle.Append("right top;")
                    Case 7 '7 справа по центру
                        strStyle.Append("right center;")
                    Case 8 '8 в правом нижнем углу
                        strStyle.Append("right bottom;")
                End Select
            End If
        End If
        If strStyle.Length > 0 Then
            hDoc.Body.Style = strStyle.ToString
            strStyle.Clear()
        End If

        'Печатаем заголовок
        Dim bCaption As String = ""
        If ReadProperty(classB, "Caption", -1, -1, bCaption, Nothing) = False Then Return "#Error"
        bCaption = UnWrapString(bCaption)
        Dim hHeader As HtmlElement = Nothing
        hHeader = hDoc.CreateElement("DIV")
        hHeader.Id = "BattleHeader"
        hHeader.SetAttribute("ClassName", "BattleCaption")
        If String.IsNullOrEmpty(bCaption) = False Then
            hHeader.InnerHtml = bCaption
        Else
            hHeader.Style = "display:none;"
        End If
        hConvas.AppendChild(hHeader)

        'печатаем VS
        Dim bVS As String = ""
        If ReadProperty(classB, "VSPicture", -1, -1, bVS, Nothing) = False Then Return "#Error"
        bVS = UnWrapString(bVS)
        If String.IsNullOrEmpty(bVS) = False Then bVS = bVS.Replace("\"c, "/"c)
        Dim imgVS As HtmlElement = hDoc.CreateElement("IMG")
        imgVS.SetAttribute("src", bVS)
        imgVS.Id = "vsPicture"
        'получаем размеры VS
        Dim vsWidth As Integer = 0, vsHeight As Integer = 0
        If ReadPropertyInt(classB, "VSPictureWidth", -1, -1, vsWidth, Nothing) = False Then Return "#Error"
        If ReadPropertyInt(classB, "VSPictureHeight", -1, -1, vsHeight, Nothing) = False Then Return "#Error"
        If vsWidth <= 0 Then vsWidth = 200
        If vsHeight <= 0 Then vsHeight = 200

        'вычисляем координаты VS
        Dim vsPos As New Point, vsPosInit As New Point
        If IsNothing(hHeader) Then
            vsPos.Y = 10
        Else
            vsPos = mapManager.GetHTMLelementCoordinates(hHeader)
            vsPos.Y += hHeader.OffsetRectangle.Height
        End If
        vsPos.X = Math.Round(((frmPlayer.wbMain.ClientRectangle.Width - vsWidth) / 2))
        vsPosInit.X = vsPos.X
        vsPosInit.Y = vsPos.Y

        imgVS.Style = "position:absolute;width:" & vsWidth.ToString & "px;height:" & vsHeight.ToString & "px;left:" & _
            vsPos.X.ToString & "px;top:" & vsPos.Y.ToString & "px;"
        hConvas.AppendChild(imgVS)

        Dim convasPos As Point = mapManager.GetHTMLelementCoordinates(hConvas)
        'создаем контейнер левых армий
        Dim hLeftArmies As HtmlElement = hDoc.CreateElement("DIV")
        hLeftArmies.SetAttribute("ClassName", "LeftArmies")
        hLeftArmies.Id = "LeftArmies"
        hConvas.AppendChild(hLeftArmies)

        'создаем контейнер правых армий
        Dim hRightArmies As HtmlElement = hDoc.CreateElement("DIV")
        hRightArmies.SetAttribute("ClassName", "RightArmies")
        hRightArmies.Id = "RightArmies"
        hConvas.AppendChild(hRightArmies)

        'выводим армии
        If PrepareFightersPosition() = "#Error" Then Return "#Error"
        For posId As Integer = 0 To lstPositions.Count - 1
            If PutArmyToBattleField(posId, False) = "#Error" Then Return "#Error"
            'Dim battlePos As FightersPosition = lstPositions(posId)
            'Dim armyId As Integer = battlePos.armyId
            'Dim fromLeft As Boolean = battlePos.fromLeft
            'Dim arrs() As String = {armyId.ToString}
            'Dim armyConvas As HtmlElement = hDoc.CreateElement("DIV")
            'armyConvas.SetAttribute("armyId", armyId.ToString)

            ''устанавливаем класс армии
            'Dim aClass As String = ""
            'If ReadProperty(classArmy, "BackgroundClass", armyId, -1, aClass, arrs) = False Then Return "#Error"
            'aClass = UnWrapString(aClass)
            'If String.IsNullOrEmpty(aClass) Then
            '    aClass = "ArmyConvas"
            'Else
            '    aClass &= " ArmyConvas"
            'End If
            'armyConvas.SetAttribute("ClassName", aClass)
            'armyConvas.SetAttribute("armyId", armyId.ToString)

            'If battlePos.fromLeft Then
            '    hLeftArmies.AppendChild(armyConvas)
            'Else
            '    hRightArmies.AppendChild(armyConvas)
            'End If

            'If armyId > -1 Then
            '    'печатаем название армии
            '    Dim aCaption As String = ""
            '    If ReadProperty(classArmy, "Caption", armyId, -1, aCaption, arrs) = False Then Return "#Error"
            '    aCaption = UnWrapString(aCaption)
            '    If String.IsNullOrEmpty(aCaption) = False Then
            '        Dim hCapt As HtmlElement = hDoc.CreateElement("H1")
            '        hCapt.InnerHtml = aCaption
            '        hCapt.SetAttribute("ClassName", "ArmyCaption")
            '        armyConvas.AppendChild(hCapt)
            '    End If

            '    Dim aCoat As String = ""
            '    If ReadProperty(classArmy, "CoatOfArms", armyId, -1, aCoat, arrs) = False Then Return "#Error"
            '    aCoat = UnWrapString(aCoat)
            '    aCoat = aCoat.Replace("\"c, "/"c)
            '    If String.IsNullOrEmpty(aCoat) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, aCoat)) Then
            '        Dim imgCoat As HtmlElement = hDoc.CreateElement("IMG")
            '        imgCoat.SetAttribute("src", aCoat)
            '        imgCoat.SetAttribute("ClassName", "ArmyCoatOfArms")
            '        armyConvas.AppendChild(imgCoat)
            '    End If
            'End If

            ''создаем таблицу для размещения армии
            'Dim hTbl As HtmlElement = hDoc.CreateElement("TABLE")
            'hTbl.SetAttribute("fromLeft", fromLeft.ToString)
            'Dim hTB As HtmlElement = hDoc.CreateElement("TBODY")
            'hTbl.AppendChild(hTB)
            'armyConvas.AppendChild(hTbl)
            'Dim TR As HtmlElement = hDoc.CreateElement("TR")
            'hTB.AppendChild(TR)

            'Dim maxPerColumn As Integer = 5
            'If ReadPropertyInt(classArmy, "FightersPerColumn", armyId, -1, maxPerColumn, arrs) = False Then Return "#Error"

            'Dim iStart As Integer, iFinish As Integer, iStep As Integer
            'If fromLeft Then
            '    iStart = battlePos.orders.Count - 1
            '    iFinish = 0
            '    iStep = -1
            'Else
            '    iStart = 0
            '    iFinish = battlePos.orders.Count - 1
            '    iStep = 1
            'End If
            'For i As Integer = iStart To iFinish Step iStep
            '    'перебираем список рангов, начиная с последнего (солдаты)
            '    Dim range As Integer = battlePos.orders.ElementAt(i).Key
            '    Dim cols As Integer = Math.Ceiling(battlePos.orders.ElementAt(i).Value.Count / maxPerColumn) 'количество ячеек, в которых надо разместить бойцов данного ранга
            '    Dim jStart As Integer, jFinish As Integer, jStep As Integer
            '    If fromLeft Then
            '        jStart = cols
            '        jFinish = 1
            '        jStep = -1
            '    Else
            '        jStart = 1
            '        jFinish = cols
            '        jStep = 1
            '    End If
            '    For j As Integer = jStart To jFinish Step jStep
            '        'создаем ячейку
            '        Dim TD As HtmlElement = hDoc.CreateElement("TD")
            '        TD.SetAttribute("Range", range.ToString) 'ранг
            '        TD.SetAttribute("ColumnInRange", j.ToString) 'номер колонки с бойцами данного ранга (надо, если все бойцы не влезли в одну колонку)
            '        TR.AppendChild(TD)
            '        'размещаем бойцов
            '        For q As Integer = (j - 1) * maxPerColumn To Math.Min(j * maxPerColumn - 1, battlePos.orders(range).Count - 1)
            '            If PutFighterToBattleField(battlePos.orders(range)(q), TD) = "#Error" Then Return "#Error"
            '        Next q
            '    Next j
            'Next i

        Next posId

        ''получаем смещение армий относительно VSPicture
        'Dim armiesOffset As Integer = 0
        'If ReadPropertyInt(classB, "ArmiesOffset", -1, -1, armiesOffset, Nothing) = False Then Return "#Error"
        'If armiesOffset <= 0 Then armiesOffset = 200

        ''смещаем контейнер левых армий
        'Dim thisPt As Point = mapManager.GetHTMLelementCoordinates(hLeftArmies)
        'Dim armiesContainerWidth As Integer = vsPosInit.X - convasPos.X - thisPt.X - armiesOffset
        'hLeftArmies.Style = "width:" & armiesContainerWidth.ToString & "px;top:" & (vsPos.Y - thisPt.Y).ToString & "px;"
        ''Опеределяем есть ли армии, ширина которых превышает ширину контейнера
        'Dim offsetToRight As Integer = 0 'при смещении контейнера левых армий вправо определяет на сколько надо сместить вправо VSPicture и контейнер правых армий
        'Dim maxWidth As Integer = armiesContainerWidth
        'For i As Integer = 0 To hLeftArmies.Children.Count - 1
        '    Dim hArmy As HtmlElement = hLeftArmies.Children(i)
        '    If hArmy.Children.Count = 0 Then Continue For
        '    hArmy = hArmy.Children(hArmy.Children.Count - 1)
        '    Dim w As Integer = hArmy.OffsetRectangle.Width
        '    If w > maxWidth Then maxWidth = w 'есть такая армия
        'Next

        'If maxWidth > armiesContainerWidth Then
        '    'есть армии, ширина которых превышает ширину контейнера
        '    offsetToRight = maxWidth - armiesContainerWidth 'устанавливаем смещение вправо
        '    Dim l As Integer = -thisPt.X - offsetToRight
        '    'меняем ширину контейнера левых армий так, чтобы влезла самая большая. При этом правый край контейнера смещается вправо на offsetToRight
        '    hLeftArmies.Style &= "width:" & maxWidth.ToString & "px;"
        '    'смещаем VSPicture
        '    vsPos.X += offsetToRight
        '    imgVS.Style &= "left:" & vsPos.X.ToString & "px;"
        'End If

        ''смещаем левые армии так, чтобы они прилегали к правому краю контейнера армий
        'thisPt = mapManager.GetHTMLelementCoordinates(hLeftArmies)
        'For i As Integer = 0 To hLeftArmies.Children.Count - 1
        '    Dim hArmy As HtmlElement = hLeftArmies.Children(i)
        '    If hArmy.Children.Count = 0 Then Continue For
        '    hArmy = hArmy.Children(hArmy.Children.Count - 1)
        '    Dim armyPos As Point = mapManager.GetHTMLelementCoordinates(hArmy)
        '    Dim l As Integer = hLeftArmies.Children(i).OffsetRectangle.Width - hArmy.OffsetRectangle.Width - armyPos.X
        '    hArmy.Style &= "position:relative;left:" & l.ToString & "px;"
        'Next

        ''смещаем контейнер правых армий
        'thisPt = mapManager.GetHTMLelementCoordinates(hRightArmies)
        'vsPos.X = vsPos.X - convasPos.X + vsWidth + armiesOffset
        'armiesContainerWidth = frmPlayer.wbMain.ClientRectangle.Width - vsPosInit.X - convasPos.X - 40
        'hRightArmies.Style = "left:" & vsPos.X.ToString & "px;width:" & armiesContainerWidth.ToString & "px;top:" & (-thisPt.Y + vsPos.Y).ToString & "px;"

        ''Опеределяем есть ли армии, ширина которых превышает ширину контейнера
        'maxWidth = armiesContainerWidth
        'For i As Integer = 0 To hRightArmies.Children.Count - 1
        '    Dim hArmy As HtmlElement = hRightArmies.Children(i)
        '    If hArmy.Children.Count = 0 Then Continue For
        '    hArmy = hArmy.Children(hArmy.Children.Count - 1)
        '    Dim w As Integer = hArmy.OffsetRectangle.Width
        '    If w > maxWidth Then maxWidth = w 'есть такая армия
        'Next
        'If maxWidth > armiesContainerWidth Then hRightArmies.Style &= "width:" & maxWidth.ToString & "px;" 'меняем ширину контейнера правых армий так, чтобы влезла самая большая

        'If offsetToRight > 0 AndAlso IsNothing(hHeader) = False Then
        '    'изменяем ширину заголовка так, чтобы он оказался посредине
        '    thisPt = mapManager.GetHTMLelementCoordinates(hRightArmies)
        '    hHeader.Style &= "position:relative;width:" & (thisPt.X + hRightArmies.OffsetRectangle.Width).ToString & "px;"
        'End If
        Dim armiesOffset As Integer = 0
        If OffsetArmiesContainers(armiesOffset) = "#Error" Then Return "#Error"

        'Выводим конвас для истории боя
        Dim hHistory As HtmlElement = hDoc.CreateElement("DIV")
        hHistory.Id = "MainConvas"
        Dim hisLeft As Integer, hisWidth As Integer, hisTop As Integer, hisOffest As Integer = 0, hisColor As String = ""
        If ReadPropertyInt(classB, "HistoryOffset", -1, -1, hisOffest, Nothing) = False Then Return "#Error"
        If ReadProperty(classB, "TextColor", -1, -1, hisColor, Nothing) = False Then Return "#Error"
        hisColor = UnWrapString(hisColor)
        If hisOffest <= 0 Then hisOffest = 300

        Dim thisPt As Point = mapManager.GetHTMLelementCoordinates(hLeftArmies)
        hisLeft = thisPt.X + hLeftArmies.OffsetRectangle.Width + 20
        hisWidth = imgVS.OffsetRectangle.Width + armiesOffset * 2 - 25
        hisTop = vsPos.Y + vsHeight + hisOffest


        Dim fightersSpeed As Integer = 0
        If ReadProperty(classB, "FightersSpeed", -1, -1, fightersSpeed, Nothing) = False Then Return "#Error"
        If fightersSpeed < 0 Then fightersSpeed = 0
        hHistory.Style = "position:absolute;left:" & hisLeft.ToString & "px;top:" & hisTop.ToString & "px;width:" & hisWidth.ToString & "px;transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;"
        If String.IsNullOrEmpty(hisColor) = False Then hHistory.Style &= "color:" & hisColor & ";"
        hHistory.InnerHtml = ""
        hDoc.Body.AppendChild(hHistory)

        'создаем footer
        Dim hFooter As HtmlElement = hDoc.CreateElement("DIV")
        hFooter.Id = "footer"
        Dim strFooter As String = "", footerVisible As Boolean = True
        If ReadProperty(classB, "Footer", -1, -1, strFooter, Nothing) = False Then Return "#Error"
        strFooter = UnWrapString(strFooter)
        hFooter.InnerHtml = strFooter
        If ReadPropertyBool(classB, "FooterVisible", -1, -1, footerVisible, Nothing) = False Then Return "#Error"
        If Not footerVisible Then hFooter.Style = "display:none"
        hDoc.Body.AppendChild(hFooter)

        Return ""
    End Function

    ''' <summary>
    ''' Выводит армию на поле боя
    ''' </summary>
    ''' <param name="ArmyPosId">Id армии в списке lstPositions. Если -1, то последняя</param>
    ''' <param name="offsetContainers">надо ли смещать контейнеры правых и левых армий, а также изменять рамеры контейнера истории боя с тем, чтобы они правильно расположились на поле боя</param>
    Public Function PutArmyToBattleField(Optional ByVal ArmyPosId As Integer = -1, Optional ByVal offsetContainers As Boolean = True, Optional armyConvas As HtmlElement = Nothing) As String
        If ArmyPosId < 0 Then ArmyPosId = lstPositions.Count - 1 'если не указана, то последняя
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        Dim hLeftArmies As HtmlElement = hDoc.GetElementById("LeftArmies")
        Dim hRightArmies As HtmlElement = hDoc.GetElementById("RightArmies")
        If IsNothing(hLeftArmies) OrElse IsNothing(hRightArmies) Then Return _ERROR("Ошибка в структуре html-документа!")
        Dim classArmy As Integer = mScript.mainClassHash("Army")
        If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
        Dim battlePos As FightersPosition = lstPositions(ArmyPosId)
        Dim armyId As Integer = battlePos.armyId
        Dim fromLeft As Boolean = battlePos.fromLeft
        Dim arrs() As String = {armyId.ToString}
        If IsNothing(armyConvas) Then
            armyConvas = hDoc.CreateElement("DIV")
            armyConvas.SetAttribute("armyId", armyId.ToString)

            'устанавливаем класс армии
            Dim aClass As String = ""
            If armyId > -1 Then
                If ReadProperty(classArmy, "BackgroundClass", armyId, -1, aClass, arrs) = False Then Return "#Error"
                aClass = UnWrapString(aClass)
                If String.IsNullOrEmpty(aClass) Then
                    aClass = "ArmyConvas"
                Else
                    aClass &= " ArmyConvas"
                End If
            Else
                aClass = "ArmyConvas"
            End If
            armyConvas.SetAttribute("ClassName", aClass)
            armyConvas.SetAttribute("armyId", armyId.ToString)

            If battlePos.fromLeft Then
                hLeftArmies.AppendChild(armyConvas)
            Else
                hRightArmies.AppendChild(armyConvas)
            End If
        Else
            armyConvas.InnerHtml = ""
        End If

        If armyId > -1 Then
            'печатаем название армии
            Dim aCaption As String = ""
            If ReadProperty(classArmy, "Caption", armyId, -1, aCaption, arrs) = False Then Return "#Error"
            aCaption = UnWrapString(aCaption)
            If String.IsNullOrEmpty(aCaption) = False Then
                Dim hCapt As HtmlElement = hDoc.CreateElement("H1")
                hCapt.InnerHtml = aCaption
                hCapt.SetAttribute("ClassName", "ArmyCaption")
                armyConvas.AppendChild(hCapt)
            End If

            Dim aCoat As String = ""
            If ReadProperty(classArmy, "CoatOfArms", armyId, -1, aCoat, arrs) = False Then Return "#Error"
            aCoat = UnWrapString(aCoat)
            aCoat = aCoat.Replace("\"c, "/"c)
            If String.IsNullOrEmpty(aCoat) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, aCoat)) Then
                Dim imgCoat As HtmlElement = hDoc.CreateElement("IMG")
                imgCoat.SetAttribute("src", aCoat)
                imgCoat.SetAttribute("ClassName", "ArmyCoatOfArms")
                armyConvas.AppendChild(imgCoat)
            End If
        End If

        'создаем таблицу для размещения армии
        Dim hTbl As HtmlElement = hDoc.CreateElement("TABLE")
        hTbl.SetAttribute("fromLeft", fromLeft.ToString)
        Dim hTB As HtmlElement = hDoc.CreateElement("TBODY")
        hTbl.AppendChild(hTB)
        armyConvas.AppendChild(hTbl)
        Dim TR As HtmlElement = hDoc.CreateElement("TR")
        hTB.AppendChild(TR)

        Dim maxPerColumn As Integer = 5
        If ReadPropertyInt(classArmy, "FightersPerColumn", armyId, -1, maxPerColumn, arrs) = False Then Return "#Error"

        Dim iStart As Integer, iFinish As Integer, iStep As Integer
        If fromLeft Then
            iStart = battlePos.orders.Count - 1
            iFinish = 0
            iStep = -1
        Else
            iStart = 0
            iFinish = battlePos.orders.Count - 1
            iStep = 1
        End If
        For i As Integer = iStart To iFinish Step iStep
            'перебираем список рангов, начиная с последнего (солдаты)
            Dim range As Integer = battlePos.orders.ElementAt(i).Key
            Dim cols As Integer = Math.Ceiling(battlePos.orders.ElementAt(i).Value.Count / maxPerColumn) 'количество ячеек, в которых надо разместить бойцов данного ранга
            Dim jStart As Integer, jFinish As Integer, jStep As Integer
            If fromLeft Then
                jStart = cols
                jFinish = 1
                jStep = -1
            Else
                jStart = 1
                jFinish = cols
                jStep = 1
            End If
            For j As Integer = jStart To jFinish Step jStep
                'создаем ячейку
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                TD.SetAttribute("Range", range.ToString) 'ранг
                TD.SetAttribute("ColumnInRange", j.ToString) 'номер колонки с бойцами данного ранга (надо, если все бойцы не влезли в одну колонку)
                TR.AppendChild(TD)
                'размещаем бойцов
                For q As Integer = (j - 1) * maxPerColumn To Math.Min(j * maxPerColumn - 1, battlePos.orders(range).Count - 1)
                    If PutFighterToBattleField(battlePos.orders(range)(q), TD) = "#Error" Then Return "#Error"
                Next q
            Next j
        Next i

        If offsetContainers Then Return OffsetArmiesContainers(Nothing) 'смещаем армии так, что бы они нормально расположились на поле боя
        Return ""
    End Function

    ''' <summary>
    ''' Функция, смещающая контейнеры правых и левых армий, а также изменяет рамеры контейнера истории боя
    ''' </summary>
    ''' <param name="armiesOffset">для получения свойства Battle.ArmiesOffset</param>
    Public Function OffsetArmiesContainers(ByRef armiesOffset As Integer) As String
        Dim classB As Integer = mScript.mainClassHash("B") ', classArmy As Integer = mScript.mainClassHash("Army")
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
        Dim hLeftArmies As HtmlElement = hDoc.GetElementById("LeftArmies")
        Dim hRightArmies As HtmlElement = hDoc.GetElementById("RightArmies")
        Dim hConvas As HtmlElement = hDoc.GetElementById("BattleConvas")
        Dim imgVS As HtmlElement = hDoc.GetElementById("vsPicture")
        If IsNothing(hLeftArmies) OrElse IsNothing(hRightArmies) OrElse IsNothing(hConvas) OrElse IsNothing(imgVS) Then Return _ERROR("Ошибка в структуре html-документа!")
        Dim hHeader As HtmlElement = hDoc.GetElementById("BattleHeader")
        Dim vsWidth As Integer = 0, vsHeight As Integer = 0
        If ReadPropertyInt(classB, "VSPictureWidth", -1, -1, vsWidth, Nothing) = False Then Return "#Error"
        If vsWidth <= 0 Then vsWidth = 200
        Dim convasPos As Point = mapManager.GetHTMLelementCoordinates(hConvas)

        'вычисляем координаты VS
        Dim vsPos As New Point, vsPosInit As New Point
        If IsNothing(hHeader) Then
            vsPos.Y = 10
        Else
            vsPos = mapManager.GetHTMLelementCoordinates(hHeader)
            vsPos.Y += hHeader.OffsetRectangle.Height
        End If
        vsPos.X = Math.Round(((frmPlayer.wbMain.ClientRectangle.Width - vsWidth) / 2))
        vsPosInit.X = vsPos.X
        vsPosInit.Y = vsPos.Y

        'получаем смещение армий относительно VSPicture
        armiesOffset = 0
        If ReadPropertyInt(classB, "ArmiesOffset", -1, -1, armiesOffset, Nothing) = False Then Return "#Error"
        If armiesOffset <= 0 Then armiesOffset = 200

        'смещаем контейнер левых армий

        Dim thisPt As Point = mapManager.GetHTMLelementCoordinates(hLeftArmies)
        Dim armiesContainerWidth As Integer = vsPosInit.X - convasPos.X - thisPt.X - armiesOffset
        hLeftArmies.Style = "width:" & armiesContainerWidth.ToString & "px;top:" & (vsPos.Y - thisPt.Y).ToString & "px;"
        'Опеределяем есть ли армии, ширина которых превышает ширину контейнера
        Dim offsetToRight As Integer = 0 'при смещении контейнера левых армий вправо определяет на сколько надо сместить вправо VSPicture и контейнер правых армий
        Dim maxWidth As Integer = armiesContainerWidth
        For i As Integer = 0 To hLeftArmies.Children.Count - 1
            Dim hArmy As HtmlElement = hLeftArmies.Children(i)
            If hArmy.Children.Count = 0 Then Continue For
            hArmy = hArmy.Children(hArmy.Children.Count - 1)
            Dim w As Integer = hArmy.OffsetRectangle.Width
            If w > maxWidth Then maxWidth = w 'есть такая армия
        Next

        If maxWidth > armiesContainerWidth Then
            'есть армии, ширина которых превышает ширину контейнера
            offsetToRight = maxWidth - armiesContainerWidth 'устанавливаем смещение вправо
            Dim l As Integer = -thisPt.X - offsetToRight
            'меняем ширину контейнера левых армий так, чтобы влезла самая большая. При этом правый край контейнера смещается вправо на offsetToRight
            hLeftArmies.Style &= "width:" & maxWidth.ToString & "px;"
            'смещаем VSPicture
            vsPos.X += offsetToRight
            imgVS.Style &= "left:" & vsPos.X.ToString & "px;"
        End If

        'смещаем левые армии так, чтобы они прилегали к правому краю контейнера армий
        thisPt = mapManager.GetHTMLelementCoordinates(hLeftArmies)
        For i As Integer = 0 To hLeftArmies.Children.Count - 1
            Dim hArmy As HtmlElement = hLeftArmies.Children(i)
            If hArmy.Children.Count = 0 Then Continue For
            hArmy = hArmy.Children(hArmy.Children.Count - 1)
            Dim armyPos As Point = mapManager.GetHTMLelementCoordinates(hArmy)
            Dim l As Integer = hLeftArmies.Children(i).OffsetRectangle.Width - hArmy.OffsetRectangle.Width - armyPos.X
            hArmy.Style &= "position:relative;left:" & l.ToString & "px;"
        Next

        'смещаем контейнер правых армий
        thisPt = mapManager.GetHTMLelementCoordinates(hRightArmies)
        vsPos.X = vsPos.X - convasPos.X + vsWidth + armiesOffset
        armiesContainerWidth = frmPlayer.wbMain.ClientRectangle.Width - vsPosInit.X - convasPos.X - 40
        hRightArmies.Style = "left:" & vsPos.X.ToString & "px;width:" & armiesContainerWidth.ToString & "px;top:" & (-thisPt.Y + vsPos.Y).ToString & "px;"

        'Опеределяем есть ли армии, ширина которых превышает ширину контейнера
        maxWidth = armiesContainerWidth
        For i As Integer = 0 To hRightArmies.Children.Count - 1
            Dim hArmy As HtmlElement = hRightArmies.Children(i)
            If hArmy.Children.Count = 0 Then Continue For
            hArmy = hArmy.Children(hArmy.Children.Count - 1)
            Dim w As Integer = hArmy.OffsetRectangle.Width
            If w > maxWidth Then maxWidth = w 'есть такая армия
        Next
        If maxWidth > armiesContainerWidth Then hRightArmies.Style &= "width:" & maxWidth.ToString & "px;" 'меняем ширину контейнера правых армий так, чтобы влезла самая большая

        If offsetToRight > 0 AndAlso IsNothing(hHeader) = False Then
            'изменяем ширину заголовка так, чтобы он оказался посредине
            thisPt = mapManager.GetHTMLelementCoordinates(hRightArmies)
            hHeader.Style &= "position:relative;width:" & (thisPt.X + hRightArmies.OffsetRectangle.Width).ToString & "px;"
        End If

        Dim hHistory As HtmlElement = hDoc.GetElementById("MainConvas")
        If IsNothing(hHistory) = False Then
            'обновляем положение и размры истории боя
            Dim fightersSpeed As Integer = 0
            If ReadProperty(classB, "FightersSpeed", -1, -1, fightersSpeed, Nothing) = False Then Return "#Error"
            If fightersSpeed < 0 Then fightersSpeed = 0
            Dim hisLeft As Integer, hisWidth As Integer, hisTop As Integer, hisOffest As Integer = 0, hisColor As String = ""
            If ReadPropertyInt(classB, "HistoryOffset", -1, -1, hisOffest, Nothing) = False Then Return "#Error"
            If ReadProperty(classB, "TextColor", -1, -1, hisColor, Nothing) = False Then Return "#Error"
            hisColor = UnWrapString(hisColor)
            If hisOffest <= 0 Then hisOffest = 300

            thisPt = mapManager.GetHTMLelementCoordinates(hLeftArmies)
            hisLeft = thisPt.X + hLeftArmies.OffsetRectangle.Width + 20
            hisWidth = imgVS.OffsetRectangle.Width + armiesOffset * 2 - 25
            hisTop = vsPos.Y + vsHeight + hisOffest
            hHistory.Style = "position:absolute;left:" & hisLeft.ToString & "px;top:" & hisTop.ToString & "px;width:" & hisWidth.ToString & "px;transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;"
            If String.IsNullOrEmpty(hisColor) = False Then hHistory.Style &= "color:" & hisColor & ";"
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Помещает бойца на поле боя
    ''' </summary>
    ''' <param name="hId">Id бойца</param>
    ''' <param name="hDest">html-элемент, внутрь которого надо поместить бойца</param>
    ''' <param name="hFighterContainer">html-контейнер бойца если он уже создан и надо изменить отображение</param>
    ''' <param name="resetPosition">Устанавливать позицию в ноль (в стандартное положение внутри армии) или использовать текущее (актуально только если боец отображен на поле боя)</param>
    Public Function PutFighterToBattleField(ByVal hId As Integer, ByRef hDest As HtmlElement, Optional ByRef hFighterContainer As HtmlElement = Nothing, Optional ByVal resetPosition As Boolean = True) As String
        Dim classB As Integer = mScript.mainClassHash("B"), pVarsUBound As Integer = -1, fPropNames() As String = Nothing, fPropClasses() As String = Nothing, _
            fPropStyle() As PropertyVisualizationStyleEnum = Nothing, hParam() As HtmlElement = Nothing, hParamTotal() As HtmlElement = Nothing, hParamInner() As HtmlElement = Nothing, _
            pValue() As Double = Nothing, pValueTotal() As Double = Nothing, textBefore() As String = Nothing, textAfter() As String = Nothing

        Dim hDoc As HtmlDocument = hDest.Document
        If IsNothing(hFighterContainer) = False Then
            hFighterContainer.InnerHtml = ""
        Else
            hFighterContainer = hDoc.CreateElement("DIV")
        End If
        hFighterContainer.SetAttribute("ClassName", "HeroContainer")
        hDest.AppendChild(hFighterContainer)
        hFighterContainer.SetAttribute("heroId", hId.ToString)
        Dim fightersSpeed As Integer = 0
        If ReadProperty(classB, "FightersSpeed", -1, -1, fightersSpeed, Nothing) = False Then Return "#Error"
        If fightersSpeed < 0 Then fightersSpeed = 0
        hFighterContainer.Style = "transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;opacity:1;"
        If resetPosition Then hFighterContainer.Style &= "left:0px;top:0px;"
        Dim hPropParent As HtmlElement = hFighterContainer


        'получаем количество выводимых свойств
        Do
            If mScript.mainClass(classB).Properties.ContainsKey("PropertyToDisplay" & (pVarsUBound + 2).ToString) Then
                pVarsUBound += 1
            Else
                Exit Do
            End If
        Loop

        If pVarsUBound > 0 Then
            'не выводится ни одно свойство - просто печатаем картинку
            ReDim fPropNames(pVarsUBound) 'имена парных свойств для отображения
            ReDim fPropClasses(pVarsUBound) 'css-классы отображения свойств
            ReDim fPropStyle(pVarsUBound) 'стили отображения свойств
            ReDim hParam(pVarsUBound) 'html-элемент для вывода текущего значения свойства (длиной 100% * Value/ValueTotal)
            ReDim hParamTotal(pVarsUBound) 'html-элемент для вывода максимального значения свойства (например, граница максимума)
            ReDim hParamInner(pVarsUBound) 'html-элемент для отображения фона
            ReDim pValue(pVarsUBound) 'текущее значение свойства
            ReDim pValueTotal(pVarsUBound) 'максимальное значение свойства
            ReDim textBefore(pVarsUBound) 'текст перед линией
            ReDim textAfter(pVarsUBound) 'текст после линии
            'первым размещается hParamTotal - он обозначает границы максимального значения. Внутри него размещается hParam. Он имеет ширину, равную отношению максимума к текущему значению,
            'а также имеет стиль overflow:hidden. Внутри помещается hParamInner, который выводит фон свойства (цвет, градиегт и т. д.). Элемент имеет ширину, равную hParamTotal, но 
            'если Value < ValueTotal, то часть его скрыта, так как у hParam overflow = hidden. Это позволяет корректно отображать градиенты - при уменьшении показателя они не сжимаются, а обрезаются.

            Dim fromLeft As Boolean = False, fromRight As Boolean = False
            For i As Integer = 0 To fPropNames.Count - 1
                textBefore(i) = ""
                textAfter(i) = ""
                'получаем имя свойства для отображения
                If ReadProperty(classB, "PropertyToDisplay" & (i + 1).ToString, -1, -1, fPropNames(i), Nothing) = False Then Return "#Error"
                fPropNames(i) = UnWrapString(fPropNames(i))
                If String.IsNullOrEmpty(fPropNames(i)) Then Continue For

                'получаем стиль отображения
                If ReadPropertyInt(classB, "PropertyToDisplay" & (i + 1).ToString & "Style", -1, -1, fPropStyle(i), Nothing) = False Then Return "#Error"
                If fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_FRONT Then
                    Dim fLeft As Boolean = CBool(hDest.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TR-TBODY-TABLE
                    If fLeft Then
                        fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                    Else
                        fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_LEFT
                    End If
                ElseIf fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_BACK Then
                    Dim fLeft As Boolean = CBool(hDest.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TR-TBODY-TABLE
                    If fLeft Then
                        fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_LEFT
                    Else
                        fPropStyle(i) = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                    End If
                End If

                'получаем текст до и после свойства
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(classB).Properties.TryGetValue("PropertyToDisplayBefore" & (i + 1).ToString, p) Then
                    If p.eventId > 0 Then
                        textBefore(i) = mScript.eventRouter.RunEvent(p.eventId, {hId.ToString}, "PropertyToDisplayBefore1", False)
                        If textBefore(i) = "#Error" Then Return textBefore(i)
                        textBefore(i) = UnWrapString(textBefore(i))
                    End If
                End If

                If mScript.mainClass(classB).Properties.TryGetValue("PropertyToDisplayAfter" & (i + 1).ToString, p) Then
                    If p.eventId > 0 Then
                        textAfter(i) = mScript.eventRouter.RunEvent(p.eventId, {hId.ToString}, "PropertyToDisplayAfter1", False)
                        If textAfter(i) = "#Error" Then Return textAfter(i)
                        textAfter(i) = UnWrapString(textAfter(i))
                    End If
                End If

                'получаем css-класс свойства
                If ReadProperty(classB, "PropertyToDisplay" & (i + 1).ToString & "Class", -1, -1, fPropClasses(i), Nothing) = False Then Return "#Error"
                fPropClasses(i) = UnWrapString(fPropClasses(i))

                If fPropNames(i) = "[Abilities]" Then
                    hParamTotal(i) = hDoc.CreateElement("DIV")
                    hParamTotal(i).SetAttribute("PropertyName", fPropNames(i))
                    hParamTotal(i).Style = "position:relative;padding:0px;margin:0px 5px 0px 5px;"
                    If PrintAbilitiesIcons(hParamTotal(i), hId, fPropStyle(i)) = "#Error" Then Return "#Error"

                    'определяем есть ли свойства слева и/или справа от картинки
                    Select Case fPropStyle(i)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            fromLeft = True
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            fromRight = True
                    End Select
                Else
                    'получаем значения свойства и ширину/высоту линии
                    Dim lineWidth As String = ""
                    If ReadFighterProperty(fPropNames(i), hId, pValue(i), {hId.ToString}) = False Then Return "#Error"
                    If pValue(i) < 0 Then pValue(i) = 0
                    If ReadFighterProperty(fPropNames(i) & "Total", hId, pValueTotal(i), {hId.ToString}) = False Then Return "#Error"

                    Dim valueDifference As Integer
                    If pValueTotal(i) <= 0 Then
                        Continue For 'свойства нет (точнее, равно 0) - пропускаем
                    Else
                        valueDifference = Math.Round(100 * pValue(i) / pValueTotal(i), 0)
                        If valueDifference > 100 Then valueDifference = 100
                    End If

                    'создаем 3 дива для вывода линии свойства
                    hParamTotal(i) = hDoc.CreateElement("DIV")
                    hParamTotal(i).SetAttribute("PropertyName", fPropNames(i))
                    hParam(i) = hDoc.CreateElement("DIV")
                    hParam(i).Style = "position: absolute;left:0px;padding:0px;margin:0px;overflow:hidden;top:0px;"
                    hParamInner(i) = hDoc.CreateElement("DIV")
                    hParam(i).AppendChild(hParamInner(i))

                    'устанавливаем класс graphProp_Total
                    hParamTotal(i).SetAttribute("className", Replace(fPropClasses(i), "graphProp_", "graphPropTotal_", 1, 1, CompareMethod.Text))

                    'устанавливаем ширину линии, обозначающую текущее значение свойства
                    If pValue(i) > pValueTotal(i) Then
                        'текущее значение больше максимальное - класс extra
                        hParamInner(i).SetAttribute("className", Replace(fPropClasses(i), "graphProp_", "graphPropExtra_", 1, 1, CompareMethod.Text))
                    Else
                        hParamInner(i).SetAttribute("ClassName", fPropClasses(i))
                    End If

                    'определяем есть ли свойства слева и/или справа от картинки
                    Select Case fPropStyle(i)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            fromLeft = True
                            lineWidth = (100 - valueDifference).ToString & "%"
                            hParam(i).Style &= "width:100%;height:100%;"
                            hParamInner(i).Style = "position:relative;top:" & lineWidth & ";" & "transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;"
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            fromRight = True
                            lineWidth = (100 - valueDifference).ToString & "%"
                            hParam(i).Style &= "width:100%;height:100%;"
                            hParamInner(i).Style = "position:relative;top:" & lineWidth & ";" & "transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;"
                        Case Else
                            lineWidth = valueDifference.ToString & "%"
                            hParam(i).Style &= "width:" & lineWidth & ";height:100%;" & "transition-property: all;transition-duration:" & fightersSpeed.ToString & "ms;"
                    End Select
                End If
            Next i

            If fromLeft AndAlso fromRight Then
                'есть вертикальные линии справа и слева
                Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                hT.SetAttribute("VerticalLines", "both")
                hFighterContainer.AppendChild(hT)
                hT.Style = "width:100%;border-spacing:0px;padding:0px"
                Dim hTB As HtmlElement = hDoc.CreateElement("TBODY")
                hT.AppendChild(hTB)
                Dim TR As HtmlElement = hDoc.CreateElement("TR")
                hTB.AppendChild(TR)
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                TR.AppendChild(TD) 'ячейка для левых свойств
                TD = hDoc.CreateElement("TD")
                TR.AppendChild(TD) 'ячейка для картинки и верхних/нижних свойств
                hPropParent = TD
                TD = hDoc.CreateElement("TD") 'ячейка для правых свойств
                TR.AppendChild(TD)
            ElseIf fromLeft Then
                'есть вертикальные линии слева
                Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                hT.SetAttribute("VerticalLines", "left")
                hFighterContainer.AppendChild(hT)
                hT.Style = "width:100%;border-spacing:0px;padding:0px"
                Dim hTB As HtmlElement = hDoc.CreateElement("TBODY")
                hT.AppendChild(hTB)
                Dim TR As HtmlElement = hDoc.CreateElement("TR")
                hTB.AppendChild(TR)
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                TR.AppendChild(TD) 'ячейка для левых свойств
                'TD.Style = "border:1px solid #000;"
                TD = hDoc.CreateElement("TD")
                TR.AppendChild(TD) 'ячейка для картинки и верхних/нижних свойств
                hPropParent = TD
            ElseIf fromRight Then
                'есть вертикальные линии справа
                Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                hT.SetAttribute("VerticalLines", "right")
                hFighterContainer.AppendChild(hT)
                hT.Style = "width:100%;border-spacing:0px;padding:0;"
                Dim hTB As HtmlElement = hDoc.CreateElement("TBODY")
                hT.AppendChild(hTB)
                Dim TR As HtmlElement = hDoc.CreateElement("TR")
                hTB.AppendChild(TR)
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                TR.AppendChild(TD) 'ячейка для картинки и верхних/нижних свойств
                hPropParent = TD
                TD = hDoc.CreateElement("TD") 'ячейка для правых свойств
                TR.AppendChild(TD)
            End If

            For i As Integer = 0 To fPropNames.Count - 1
                'выводим свойства над картинкой
                If IsNothing(hParamTotal(i)) Then Continue For
                Select Case fPropStyle(i)
                    Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP
                        If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                            'text before line is absend
                            hPropParent.AppendChild(hParamTotal(i))
                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                        Else
                            'text before or/and after line is exists
                            Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                            hT.SetAttribute("Align", "center")
                            hT.Style = "width:100%;border-spacing: 0px;padding:0px"
                            hPropParent.AppendChild(hT)
                            Dim hB As HtmlElement = hDoc.CreateElement("TBODY")
                            hT.AppendChild(hB)
                            Dim TR As HtmlElement = hDoc.CreateElement("TR")
                            hB.AppendChild(TR)
                            Dim TD As HtmlElement

                            'textBefore
                            If textBefore(i).Length > 0 Then
                                TD = hDoc.CreateElement("TD")
                                TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                TD.InnerHtml = textBefore(i)
                                TD.SetAttribute("text", "before")
                                TR.AppendChild(TD)
                            End If

                            'property line
                            TD = hDoc.CreateElement("TD")
                            TD.Style = "width:100%"
                            TR.AppendChild(TD)
                            TD.AppendChild(hParamTotal(i))

                            'textAfter
                            If textAfter(i).Length > 0 Then
                                TD = hDoc.CreateElement("TD")
                                TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                TD.InnerHtml = textAfter(i)
                                TD.SetAttribute("text", "after")
                                TR.AppendChild(TD)
                            End If

                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                        End If
                    Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                        If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                            'text before line is absend
                            '.. TD(1)-TR-TD(0)
                            Dim hEL As HtmlElement = hPropParent.Parent.Children(0)
                            hEL.AppendChild(hParamTotal(i))
                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                        Else
                            'text before or/and after line is exists
                            Dim hEL As HtmlElement = hPropParent.Parent.Children(0)
                            Dim txtRange As Integer = 1
                            If textBefore(i).Length > 0 Then txtRange += 1
                            If textAfter(i).Length > 0 Then txtRange += 1

                            If hEL.Children.Count > 0 AndAlso hEL.Children(0).TagName = "TABLE" AndAlso hEL.Children(0).Children(0).Children.Count = txtRange Then
                                'таблица хотя бы с одним свойством в данной позиции уже есть. Также в таблице нужное количество рядов. Добавляем в эту же таблицу новые ячейки
                                Dim TR As HtmlElement, TD As HtmlElement

                                'text before
                                If textBefore(i).Length > 0 Then
                                    TR = hEL.Children(0).Children(0).Children(0) 'TABLE-TBODY-TR(0)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TR.AppendChild(TD)
                                    TD.InnerHtml = textBefore(i)
                                    TD.SetAttribute("text", "before")
                                End If

                                'property line
                                TR = hEL.Children(0).Children(0).Children(1) 'TABLE-TBODY-TR(1)
                                TD = hDoc.CreateElement("TD")
                                TD.AppendChild(hParamTotal(i))
                                TR.AppendChild(TD)

                                'text after
                                If textAfter(i).Length > 0 Then
                                    TR = hEL.Children(0).Children(0).Children(2) 'TABLE-TBODY-TR(2)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TR.AppendChild(TD)
                                    TD.InnerHtml = textAfter(i)
                                    TD.SetAttribute("text", "after")
                                End If
                            Else
                                'в данной позиции подобного свойства пока нет - это первое (или количество рядов отличное)
                                Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                                hT.Style = "height:100%;border-spacing: 0px;padding:0px;"
                                hEL.AppendChild(hT)
                                Dim hB As HtmlElement = hDoc.CreateElement("TBODY")
                                hT.AppendChild(hB)
                                Dim TR As HtmlElement, TD As HtmlElement

                                'text before
                                If textBefore(i).Length > 0 Then
                                    TR = hDoc.CreateElement("TR")
                                    hB.AppendChild(TR)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textBefore(i)
                                    TD.SetAttribute("text", "before")
                                    TR.AppendChild(TD)
                                End If

                                'property line
                                TR = hDoc.CreateElement("TR")
                                hB.AppendChild(TR)
                                TR.Style = "height:100%;"
                                TD = hDoc.CreateElement("TD")
                                TD.AppendChild(hParamTotal(i))
                                TR.AppendChild(TD)

                                'text after
                                If textAfter(i).Length > 0 Then
                                    TR = hDoc.CreateElement("TR")
                                    hB.AppendChild(TR)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textAfter(i)
                                    TD.SetAttribute("text", "after")
                                    TR.AppendChild(TD)
                                End If
                            End If

                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                            hParamTotal(i).Style = "height:100%;"
                        End If
                    Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                        If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                            'text before line is absend
                            Dim hEL As HtmlElement
                            If fromLeft Then
                                hEL = hPropParent.Parent.Children(2)
                            Else
                                hEL = hPropParent.Parent.Children(1)
                            End If
                            hEL.AppendChild(hParamTotal(i))
                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                        Else
                            'text before or/and after line is exists
                            Dim hEL As HtmlElement
                            If fromLeft Then
                                hEL = hPropParent.Parent.Children(2)
                            Else
                                hEL = hPropParent.Parent.Children(1)
                            End If
                            Dim txtRange As Integer = 1
                            If textBefore(i).Length > 0 Then txtRange += 1
                            If textAfter(i).Length > 0 Then txtRange += 1

                            If hEL.Children.Count > 0 AndAlso hEL.Children(0).TagName = "TABLE" AndAlso hEL.Children(0).Children(0).Children.Count = txtRange Then
                                'таблица хотя бы с одним свойством в данной позиции уже есть. Также в таблице нужное количество рядов. Добавляем в эту же таблицу новые ячейки
                                Dim TR As HtmlElement, TD As HtmlElement

                                'text before
                                If textBefore(i).Length > 0 Then
                                    TR = hEL.Children(0).Children(0).Children(0) 'TABLE-TBODY-TR(0)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textBefore(i)
                                    TD.SetAttribute("text", "before")
                                    TR.AppendChild(TD)
                                End If

                                'property line
                                TR = hEL.Children(0).Children(0).Children(1) 'TABLE-TBODY-TR(1)
                                TD = hDoc.CreateElement("TD")
                                TD.AppendChild(hParamTotal(i))
                                TR.AppendChild(TD)

                                'text after
                                If textAfter(i).Length > 0 Then
                                    TR = hEL.Children(0).Children(0).Children(2) 'TABLE-TBODY-TR(2)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textAfter(i)
                                    TD.SetAttribute("text", "after")
                                    TR.AppendChild(TD)
                                End If
                            Else
                                Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                                hT.Style = "height:100%;border-spacing: 0px;padding:0px"
                                hEL.AppendChild(hT)
                                Dim hB As HtmlElement = hDoc.CreateElement("TBODY")
                                hT.AppendChild(hB)
                                Dim TR As HtmlElement, TD As HtmlElement

                                'text before
                                If textBefore(i).Length > 0 Then
                                    TR = hDoc.CreateElement("TR")
                                    hB.AppendChild(TR)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textBefore(i)
                                    TD.SetAttribute("text", "before")
                                    TR.AppendChild(TD)
                                End If

                                'property line
                                TR = hDoc.CreateElement("TR")
                                hB.AppendChild(TR)
                                TR.Style = "height:100%;"
                                TD = hDoc.CreateElement("TD")
                                TD.AppendChild(hParamTotal(i))
                                TR.AppendChild(TD)

                                'text after
                                If textAfter(i).Length > 0 Then
                                    TR = hDoc.CreateElement("TR")
                                    hB.AppendChild(TR)
                                    TD = hDoc.CreateElement("TD")
                                    TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                                    TD.InnerHtml = textAfter(i)
                                    TD.SetAttribute("text", "after")
                                    TR.AppendChild(TD)
                                End If
                            End If

                            If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                            hParamTotal(i).Style = "height:100%;"
                        End If
                End Select
            Next i
        End If

        'создаем картинку персонажа
        Dim hImg As HtmlElement = hDoc.CreateElement("IMG")
        Dim fPicture As String = "", fPicWidth As String = "", fPicHeight As String = "", heroArr() As String = {hId.ToString}
        If ReadFighterProperty("Picture", hId, fPicture, heroArr) = False Then Return "#Error"
        fPicture = UnWrapString(fPicture)
        If String.IsNullOrEmpty(fPicture) = False Then fPicture = fPicture.Replace("\"c, "/"c)
        hImg.SetAttribute("src", fPicture)
        hImg.SetAttribute("imageType", "hero")

        If ReadFighterProperty("PicWidth", hId, fPicWidth, heroArr) = False Then Return "#Error"
        fPicWidth = UnWrapString(fPicWidth)
        If ReadFighterProperty("PicHeight", hId, fPicHeight, heroArr) = False Then Return "#Error"
        fPicHeight = UnWrapString(fPicHeight)
        Dim strStyle As String = ""
        If String.IsNullOrEmpty(fPicWidth) = False Then
            If IsNumeric(fPicWidth) Then fPicWidth &= "px"
            strStyle = "width:" & fPicWidth & ";"
        End If
        If String.IsNullOrEmpty(fPicHeight) = False Then
            If IsNumeric(fPicHeight) Then fPicHeight &= "px"
            strStyle &= "height:" & fPicHeight & ";"
        End If
        If String.IsNullOrEmpty(strStyle) = False Then hImg.Style = strStyle

        'hImg.SetAttribute("ClassName", "HeroPicture")
        hPropParent.AppendChild(hImg)

        Dim template As String = ""
        template = PropertiesRouter(mScript.mainClassHash("H"), "DescriptionTemplate", heroArr, Nothing, PropertiesOperationEnum.PROPERTY_GET)
        If template = "#Error" Then Return template
        template = UnWrapString(template)

        Dim hTTip As HtmlElement = hDoc.CreateElement("SPAN")
        hTTip.SetAttribute("ClassName", "tooltiptext")
        hTTip.InnerHtml = template
        hFighterContainer.AppendChild(hTTip)
        AddHandler hFighterContainer.MouseEnter, Sub(sender As HtmlElement, e As HtmlElementEventArgs)
                                                     template = PropertiesRouter(mScript.mainClassHash("H"), "DescriptionTemplate", heroArr, Nothing, PropertiesOperationEnum.PROPERTY_GET)
                                                     If template = "#Error" Then Return
                                                     template = UnWrapString(template)
                                                     hTTip.InnerHtml = template
                                                 End Sub

        'создаем картинку для эффектов
        Dim hEffect As HtmlElement = hDoc.CreateElement("IMG")
        hEffect.Style = "display:none;position:absolute;left:" & hImg.OffsetRectangle.X.ToString & "px;top:" & hImg.OffsetRectangle.Y.ToString & "px;width:" & hImg.OffsetRectangle.Width.ToString & "px;height:" & hImg.OffsetRectangle.Height.ToString & "px;"
        hEffect.SetAttribute("ClassName", "EffectPicture")
        hEffect.SetAttribute("imageType", "effect")
        hPropParent.AppendChild(hEffect)

        If pVarsUBound > -1 Then
            For i As Integer = 0 To fPropNames.Count - 1
                'выводим свойства под картинкой
                If IsNothing(hParamTotal(i)) OrElse fPropStyle(i) <> PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM Then Continue For
                If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                    'текста до и после линии нет
                    hPropParent.AppendChild(hParamTotal(i))
                    If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                Else
                    'есть текст до и/или после линии
                    Dim hT As HtmlElement = hDoc.CreateElement("TABLE")
                    hT.SetAttribute("Align", "center")
                    hT.Style = "width:100%;border-spacing: 0px;padding:0px"
                    hPropParent.AppendChild(hT)
                    Dim hB As HtmlElement = hDoc.CreateElement("TBODY")
                    hT.AppendChild(hB)
                    Dim TR As HtmlElement, TD As HtmlElement
                    TR = hDoc.CreateElement("TR")
                    hB.AppendChild(TR)

                    'textBefore
                    If textBefore(i).Length > 0 Then
                        TD = hDoc.CreateElement("TD")
                        TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                        TD.InnerHtml = textBefore(i)
                        TD.SetAttribute("text", "before")
                        TR.AppendChild(TD)
                    End If

                    'line
                    hB.AppendChild(TR)
                    TD = hDoc.CreateElement("TD")
                    TD.Style = "width:100%"
                    TR.AppendChild(TD)
                    TD.AppendChild(hParamTotal(i))

                    'textAfter
                    If textAfter(i).Length > 0 Then
                        TD = hDoc.CreateElement("TD")
                        TD.SetAttribute("ClassName", Replace(fPropClasses(i), "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                        TD.InnerHtml = textAfter(i)
                        TD.SetAttribute("text", "after")
                        TR.AppendChild(TD)
                    End If

                    If IsNothing(hParam(i)) = False Then hParamTotal(i).AppendChild(hParam(i))
                End If
            Next i

            'получаем размеры картинки и ее родительского элемента
            Dim picPos As Point = mapManager.GetHTMLelementCoordinates(hImg)
            Dim divPos As Point = mapManager.GetHTMLelementCoordinates(hPropParent)

            For i As Integer = 0 To hParamTotal.Count - 1
                'устанавливаем ширину / высоту линий, которые зависят от размеров картинок
                If IsNothing(hParamTotal(i)) Then Continue For
                If pValueTotal(i) > 0 Then hParamTotal(i).SetAttribute("Title", Math.Round(100 * pValue(i) / pValueTotal(i), 0).ToString & "% (" & pValue(i).ToString & "/" & pValueTotal(i).ToString & ")")
                Select Case fPropStyle(i)
                    Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                        If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                            hParamTotal(i).Style = "left:" & (picPos.X - divPos.X).ToString & "px;width:" & hImg.OffsetRectangle.Width.ToString & "px;"
                            If IsNothing(hParamInner(i)) = False Then hParamInner(i).Style &= "width:" & hImg.OffsetRectangle.Width.ToString & "px;"
                        Else
                            '..-TD-TR-TB-TABLE
                            Dim hEL As HtmlElement = hParamTotal(i).Parent.Parent.Parent.Parent
                            hEL.Style = "left:" & (picPos.X - divPos.X).ToString & "px;width:" & hImg.OffsetRectangle.Width.ToString & "px;"
                            If IsNothing(hParamInner(i)) = False Then hParamInner(i).Style &= "width:" & hImg.OffsetRectangle.Width.ToString & "px;"
                        End If
                    Case PropertyVisualizationStyleEnum.VERTICAL_LEFT, PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                        If textBefore(i).Length = 0 AndAlso textAfter(i).Length = 0 Then
                            hParamTotal(i).Style = "top:" & (picPos.Y - divPos.Y).ToString & "px;height:" & hImg.OffsetRectangle.Height.ToString & "px;display:inline-block;"
                            If IsNothing(hParamInner(i)) = False Then hParamInner(i).Style &= "height:" & hImg.OffsetRectangle.Height.ToString & "px;"
                        Else
                            '..-TD-TR-TB-TABLE
                            Dim hEL As HtmlElement = hParamTotal(i).Parent.Parent.Parent.Parent
                            hEL.Style &= "top:" & (picPos.Y - divPos.Y).ToString & "px;height:" & hImg.OffsetRectangle.Height.ToString & "px;"
                            If IsNothing(hParamInner(i)) = False Then hParamInner(i).Style &= "height:" & hImg.OffsetRectangle.Height.ToString & "px;"
                        End If
                End Select
            Next i
        End If

        'смещаем бойца за пределы поля боя (так, чтобы он затем какбы вылетел откуда-то)
        Dim fPos As Point = GetFighterPositionOutOfBattlefield(hFighterContainer)
        If Object.Equals(fPos, Point.Empty) Then Return "#Error"
        hFighterContainer.Style &= "left:" & fPos.X.ToString & "px;top:" & fPos.Y.ToString & "px;"
        lsthFightersToMoveInto.Add(hFighterContainer)
        timMoveFighters.Enabled = True
        Return ""
    End Function

    ''' <summary>
    ''' возвращает случайные координаты бойца за пределами поля боя
    ''' </summary>
    ''' <param name="hFighter">html-контейнер бойца</param>
    Private Function GetFighterPositionOutOfBattlefield(ByVal hFighter As HtmlElement) As Point
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then
            mScript.LAST_ERROR = "Не найден html-документ главного окна!"
            Return Point.Empty
        End If
        Dim imgVS As HtmlElement = hDoc.GetElementById("vsPicture")
        If IsNothing(imgVS) Then
            mScript.LAST_ERROR = "Ошибка в структуре html-документа!"
            Return Point.Empty
        End If

        Dim toLeft As Boolean = CBool(hFighter.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..TD-TR-TBODY-TABLE
        Dim fSize As Size = hFighter.OffsetRectangle.Size 'размер контейнера с бойцом
        Dim fPos As Point = mapManager.GetHTMLelementCoordinates(hFighter)
        Dim direction As Double = Rnd() * 3 '<1 - сверху, 1-2 - сбоку, >2 - снизу
        'координаты центра VS по Х
        Dim vsCenter As Integer = Math.Round(((frmPlayer.wbMain.ClientRectangle.Width - imgVS.OffsetRectangle.Width) / 2))

        'вычисляем координаты бойца за пределами поля боя
        Dim newPos As New Point
        If toLeft Then
            'боец в левой армии - должен появиться с левой части экрана
            Select Case direction
                Case Is < 1 'боец появляется сверху
                    newPos.Y = -fPos.Y - fSize.Height - 20
                    Dim startX As Integer = -fPos.X - fSize.Width
                    Dim endX As Integer = startX + vsCenter
                    newPos.X = Math.Round(Rnd() * (endX - startX) + startX)
                Case 1 To 2 'боец появляется слева
                    newPos.X = -fPos.X - fSize.Width - 20
                    Dim startY As Integer = -fPos.Y
                    Dim endY As Integer = -fPos.Y - fSize.Height + frmPlayer.wbMain.ClientRectangle.Height
                    newPos.Y = Math.Round(Rnd() * (endY - startY) + startY)
                Case Else 'боец появляется снизу
                    newPos.Y = -fPos.Y + frmPlayer.wbMain.ClientRectangle.Height + 20
                    Dim startX As Integer = -fPos.X - fSize.Width
                    Dim endX As Integer = startX + vsCenter
                    newPos.X = Math.Round(Rnd() * (endX - startX) + startX)
            End Select
        Else
            'боец в правой армии - должен появиться с правой части экрана
            Select Case direction
                Case Is < 1 'боец появляется сверху
                    newPos.Y = -fPos.Y - fSize.Height - 20
                    Dim startX As Integer = -fPos.X + vsCenter
                    Dim endX As Integer = -fPos.X + frmPlayer.wbMain.ClientRectangle.Width
                    newPos.X = Math.Round(Rnd() * (endX - startX) + startX)
                Case 1 To 2 'боец появляется справа
                    newPos.X = -fPos.X + frmPlayer.wbMain.ClientRectangle.Width + 20
                    Dim startY As Integer = -fPos.Y
                    Dim endY As Integer = -fPos.Y - fSize.Height + frmPlayer.wbMain.ClientRectangle.Height
                    newPos.Y = Math.Round(Rnd() * (endY - startY) + startY)
                Case Else 'боец появляется снизу
                    newPos.Y = -fPos.Y + frmPlayer.wbMain.ClientRectangle.Height + 20
                    Dim startX As Integer = -fPos.X + vsCenter
                    Dim endX As Integer = -fPos.X + frmPlayer.wbMain.ClientRectangle.Width
                    newPos.X = Math.Round(Rnd() * (endX - startX) + startX)
            End Select
        End If

        Return newPos
    End Function

    ''' <summary>
    ''' Выдвигает бойца на передовую
    ''' </summary>
    ''' <param name="hDoc">html-документ</param>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="toLeft">при autoPosition = False указывает помещать бойца слева или справа; при autoPosition = True служит для получения позиции</param>
    ''' <param name="autoPosition">Определить значение toLeft автоматически или использовать заданное значение</param>
    Private Function MoveFighterToFront(ByRef hDoc As HtmlDocument, ByVal fId As Integer, ByRef toLeft As Boolean, Optional autoPosition As Boolean = False) As String
        Do
            If lsthFightersToMoveInto.Count > 0 OrElse lstFightersToMoveOut.Count > 0 Then
                Application.DoEvents()
            Else
                Exit Do
            End If
        Loop

        'получаем html-контейнер бойца
        Dim hFighter As HtmlElement = FindFighterHTMLContainer(fId, hDoc)
        If IsNothing(hFighter) Then Return _ERROR("Не найден html-контейнер бойца с Id " & fId.ToString)

        If autoPosition Then
            toLeft = CBool(hFighter.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..TD-TR-TBODY-TABLE
        End If

        'получаем координаты vsPicture в позиции middle bottom
        Dim vsImg As HtmlElement = hDoc.GetElementById("vsPicture")
        If IsNothing(vsImg) Then Return _ERROR("Ошибка в структуре html-документа. Картинка vsPicture не найдена.")
        Dim vsPos As Point = mapManager.GetHTMLelementCoordinates(vsImg)
        vsPos.X += vsImg.OffsetRectangle.Width / 2
        vsPos.Y += vsImg.OffsetRectangle.Height

        'вычисляем координаты для размещения бойца
        Dim fPos As Point = mapManager.GetHTMLelementCoordinates(hFighter)
        fPos.Y = vsPos.Y - fPos.Y
        If toLeft Then
            'разместить бойца слева
            fPos.X = vsPos.X - hFighter.OffsetRectangle.Width - fPos.X
        Else
            'разместить бойца справа
            fPos.X = vsPos.X - fPos.X
        End If
        hFighter.Style &= "left:" & fPos.X.ToString & "px;top:" & fPos.Y.ToString & "px;"
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает бойца из передовой обратно
    ''' </summary>
    ''' <param name="hDoc">html-документ</param>
    ''' <param name="fId">Id бойца</param>
    Public Function MoveFighterBack(ByRef hDoc As HtmlDocument, ByVal fId As Integer) As String
        Do
            If lsthFightersToMoveInto.Count > 0 OrElse lstFightersToMoveOut.Count > 0 Then
                Application.DoEvents()
            Else
                Exit Do
            End If
        Loop

        Dim hFighter As HtmlElement = FindFighterHTMLContainer(fId, hDoc)
        If IsNothing(hFighter) Then Return _ERROR("Не найден html-контейнер бойца с Id " & fId.ToString)

        hFighter.Style &= "left:0px;top:0px;"
        Return ""
    End Function

    ''' <summary>
    ''' Добавляет бойца во время боя, когда уже первичное заполнение поля боя армиями закончено
    ''' </summary>
    ''' <param name="fId">Id добавляемого бойца</param>
    Public Function AppendFighterInBattle(ByVal fId As Integer) As String
        'получаем html-документ
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return "Не найден html-документ главного окна!"
        'обновляем структуру с расположением бойцов
        If PrepareFightersPosition() = "#Error" Then Return "#Error"

        'получаем координаты в структуре lstPositions добавляемого бойца
        Dim toLeft As Boolean
        Dim armyId As Integer = Fighters(fId).armyId
        Dim lstPos As Integer = -1, order As Integer = -1, fPos As Integer = -1
        For i As Integer = 0 To lstPositions.Count - 1
            If lstPositions(i).armyId = armyId Then
                lstPos = i
                For j As Integer = 0 To lstPositions(i).orders.Count - 1
                    Dim lst As List(Of Integer) = lstPositions(i).orders.ElementAt(j).Value
                    For u As Integer = 0 To lst.Count - 1
                        If lst(u) = fId Then
                            order = lstPositions(i).orders.ElementAt(j).Key
                            fPos = u
                            toLeft = lstPositions(i).fromLeft
                            Exit For
                        End If
                    Next u
                    If fPos > -1 Then Exit For
                Next j
                If fPos > -1 Then Exit For
            End If
        Next i
        If fPos < 0 Then Return _ERROR("Ошибка при удалении бойца с поля боя: не найдено положение бойца в армии!")

        Dim classArmy As Integer = mScript.mainClassHash("Army")
        Dim FightersPerColumn As Integer = 4
        If ReadPropertyInt(classArmy, "FightersPerColumn", armyId, -1, FightersPerColumn, {armyId.ToString}) = False Then Return "#Error"
        If lstPositions(lstPos).orders.Count = 1 AndAlso lstPositions(lstPos).orders(order).Count = 1 Then
            'первый боец в армии - создаем новую армию
            If PutArmyToBattleField(lstPos) = "#Error" Then Return "#Error"
        ElseIf lstPositions(lstPos).orders(order).Count = 1 OrElse _
            Math.Floor((lstPositions(lstPos).orders(order).Count - 1) / FightersPerColumn) = (lstPositions(lstPos).orders(order).Count - 1) / FightersPerColumn Then
            'A) первый боец данного ранга - создаем новую ячейку
            'B) ячейки с бойцами данного ранга заполнены - создаем новую ячейку
            Dim newColumnNumber As Integer = Math.Ceiling(lstPositions(lstPos).orders(order).Count / FightersPerColumn)

            'получаем html-элемент контейнера армии
            Dim ArmiesConvas As HtmlElement
            If toLeft Then
                ArmiesConvas = hDoc.GetElementById("LeftArmies")
            Else
                ArmiesConvas = hDoc.GetElementById("RightArmies")
            End If
            If IsNothing(ArmiesConvas) Then Return _ERROR("Не удалось разместить бойца на поле боя. Ошибка в структуре html-документа.")
            Dim TR As HtmlElement = Nothing
            For i As Integer = 0 To ArmiesConvas.Children.Count - 1
                'перебираем все армии в контейнере
                Dim hArmy As HtmlElement = ArmiesConvas.Children(i)
                Dim army As String = hArmy.GetAttribute("armyId")
                If String.IsNullOrEmpty(army) OrElse Val(army) <> armyId Then Continue For 'если армия персонажа другая - пропускаем армию

                'контейнер армии найден. В нем TABLE-TBODY-TR-TD(x). А в начале может быть подпись и герб
                Dim hTable As HtmlElement = hArmy.Children(hArmy.Children.Count - 1)
                If hTable.TagName <> "TABLE" Then Continue For

                TR = hTable.Children(0).Children(0)
            Next i
            If IsNothing(TR) Then Return _ERROR("Не удалось разместить бойца на поле боя. Ошибка в структуре html-документа.")


            Dim nextTD As HtmlElement = Nothing
            If toLeft Then
                'армия расположена слева
                For i As Integer = 0 To TR.Children.Count - 1
                    'перебираем все ячейки таблицы, каждая из которых содержит бойцов определенного ранга
                    'ищем ячейку с рангом более низким, чем у добавляемого бойца. Если найдем - то перед ней надо вставить ячейку для нашего бойца
                    Dim TD As HtmlElement = TR.Children(i) 'TD(Range, ColumnInRange)
                    Dim strRange As String = TD.GetAttribute("Range")
                    If String.IsNullOrEmpty(strRange) = False AndAlso Val(strRange) < order Then
                        nextTD = TD 'такая ячейка найдена
                        Exit For
                    End If
                Next i
            Else
                'армия расположена справа
                For i As Integer = 0 To TR.Children.Count - 1
                    'перебираем все ячейки таблицы, каждая из которых содержит бойцов определенного ранга
                    'ищем ячейку с рангом более высоким, чем у добавляемого бойца. Если найдем - то перед ней надо вставить ячейку для нашего бойца
                    Dim TD As HtmlElement = TR.Children(i) 'TD(Range, ColumnInRange)
                    Dim strRange As String = TD.GetAttribute("Range")
                    If String.IsNullOrEmpty(strRange) = False AndAlso Val(strRange) > order Then
                        nextTD = TD 'такая ячейка найдена
                        Exit For
                    End If
                Next i
            End If

            'создаем новую ячейку
            Dim newTD As HtmlElement = hDoc.CreateElement("TD")
            newTD.SetAttribute("Range", order.ToString)
            newTD.SetAttribute("ColumnInRange", newColumnNumber.ToString)
            If IsNothing(nextTD) Then
                'ячейка с более низким значеним ранга, чем у нового бойца, не существует - добавляем новую ячейку в конец
                TR.AppendChild(newTD)
            Else
                'ячейка с более низким рангом найдена - добавляем перед ней
                nextTD.InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, newTD)
            End If

            'добавляем бойца в новую ячейку
            If PutFighterToBattleField(fId, newTD) = "#Error" Then Return "#Error"
        Else
            'существует незаполненная до конца ячейка с бойцами данного ранга - помещаем бойца в нее

            'получаем html-элемент контейнера армии
            Dim ArmiesConvas As HtmlElement
            If toLeft Then
                ArmiesConvas = hDoc.GetElementById("LeftArmies")
            Else
                ArmiesConvas = hDoc.GetElementById("RightArmies")
            End If
            If IsNothing(ArmiesConvas) Then Return _ERROR("Не удалось разместить бойца на поле боя. Ошибка в структуре html-документа.")
            Dim TR As HtmlElement = Nothing
            For i As Integer = 0 To ArmiesConvas.Children.Count - 1
                'перебираем все армии в контейнере
                Dim hArmy As HtmlElement = ArmiesConvas.Children(i)
                Dim army As String = hArmy.GetAttribute("armyId")
                If String.IsNullOrEmpty(army) OrElse Val(army) <> armyId Then Continue For 'если армия персонажа другая - пропускаем армию

                'контейнер армии найден. В нем TABLE-TBODY-TR-TD(x). А в начале может быть подпись и герб
                Dim hTable As HtmlElement = hArmy.Children(hArmy.Children.Count - 1)
                If hTable.TagName <> "TABLE" Then Continue For

                TR = hTable.Children(0).Children(0)
            Next i
            If IsNothing(TR) Then Return _ERROR("Не удалось разместить бойца на поле боя. Ошибка в структуре html-документа.")

            Dim curTD As HtmlElement = Nothing
            If toLeft Then
                'армия расположена слева
                For i As Integer = 0 To TR.Children.Count - 1
                    'перебираем все ячейки таблицы, каждая из которых содержит бойцов определенного ранга
                    'ищем ячейку с рангом, равным рангу текущего бойца
                    Dim TD As HtmlElement = TR.Children(i) 'TD(Range, ColumnInRange)
                    Dim strRange As String = TD.GetAttribute("Range")
                    If String.IsNullOrEmpty(strRange) = False AndAlso Val(strRange) = order Then
                        curTD = TD 'такая ячейка найдена
                        Exit For
                    End If
                Next i
            Else
                'армия расположена справа
                For i As Integer = TR.Children.Count - 1 To 0 Step -1
                    'перебираем все ячейки таблицы, каждая из которых содержит бойцов определенного ранга
                    'ищем ячейку с рангом, равным рангу текущего бойца
                    Dim TD As HtmlElement = TR.Children(i) 'TD(Range, ColumnInRange)
                    Dim strRange As String = TD.GetAttribute("Range")
                    If String.IsNullOrEmpty(strRange) = False AndAlso Val(strRange) = order Then
                        curTD = TD 'такая ячейка найдена
                        Exit For
                    End If
                Next i
            End If
            If IsNothing(curTD) Then Return _ERROR("Не удалось разместить бойца на поле боя. Ошибка в структуре html-документа.")
            'добавляем бойца в полученную ячейку
            If PutFighterToBattleField(fId, curTD) = "#Error" Then Return "#Error"
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Функция удаляет указанного бойца с поля боя
    ''' </summary>
    ''' <param name="fId"></param>
    Public Function RemoveFighterFromBattlefield(ByVal fId As Integer) As String
        'получаем html-документ
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")

        'получаем html-контейнер удаляемого бойца
        Dim hFighter = FindFighterHTMLContainer(fId, hDoc)
        If IsNothing(hFighter) Then Return ""

        'получаем координаты в структуре lstPositions удаляемого бойца
        Dim toLeft As Boolean = CBool(hFighter.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..TD(Range, ColumnInRange)-TR-TBODY-TABLE(fromLeft)
        Dim armyId As Integer = Fighters(fId).armyId
        Dim lstPos As Integer = -1, order As Integer = -1, fPos As Integer = -1
        For i As Integer = 0 To lstPositions.Count - 1
            If lstPositions(i).armyId = armyId AndAlso lstPositions(i).fromLeft = toLeft Then
                lstPos = i
                For j As Integer = 0 To lstPositions(i).orders.Count - 1
                    Dim lst As List(Of Integer) = lstPositions(i).orders.ElementAt(j).Value
                    For u As Integer = 0 To lst.Count - 1
                        If lst(u) = fId Then
                            order = lstPositions(i).orders.ElementAt(j).Key
                            fPos = u
                            Exit For
                        End If
                    Next u
                    If fPos > -1 Then Exit For
                Next j
                Exit For
            End If
        Next i
        If fPos < 0 Then Return "" '_ERROR("Ошибка при удалении бойца с поля боя: не найдено положение бойца в армии!")

        Dim classArmy As Integer = mScript.mainClassHash("Army")
        Dim FightersPerColumn As Integer = 4
        If ReadPropertyInt(classArmy, "FightersPerColumn", armyId, -1, FightersPerColumn, {armyId.ToString}) = False Then Return "#Error"
        If lstPositions(lstPos).orders.Count = 1 AndAlso lstPositions(lstPos).orders(order).Count = 1 Then
            'последний боец в армии - удаляем всю армию
            Dim msArmy As mshtml.IHTMLDOMNode = hFighter.Parent.Parent.Parent.Parent.Parent.DomElement
            msArmy.removeNode(True)
        ElseIf lstPositions(lstPos).orders(order).Count = 1 Then
            'последний боец данного ранга - удаляем ячейку
            Dim msTD As mshtml.IHTMLDOMNode = hFighter.Parent.DomElement
            msTD.removeNode(True)
        ElseIf lstPositions(lstPos).orders(order).Count <= FightersPerColumn OrElse fPos >= Math.Floor((lstPositions(lstPos).orders(order).Count - 1) / FightersPerColumn) * FightersPerColumn Then
            'A) все бойцы данного ранга находятся в одной ячейке (и удаляемый боец - не последний) - удаляем бойца
            'B) бойцы расположены в нескольких ячейках, но удаляем из последней ячейки - удаляем бойца или ячейку
            Dim hPar As HtmlElement = hFighter.Parent
            If hPar.Children.Count = 1 Then
                'в данной ячейке только 1 боец - удаляем ячейку
                Dim msPar As mshtml.IHTMLDOMNode = hPar.DomElement
                msPar.removeNode(True)
            Else
                'удаляем бойца
                Dim msFighter As mshtml.IHTMLDOMNode = hFighter.DomElement
                msFighter.removeNode(True)
            End If
        Else
            'бойцы данного ранга располагаются в нескольких ячейках, при этом боец удаляется не из последней.
            'надо удалить бойца, а на его место поставить последнего

            'получаем бойца, который станет на замену удаляемому
            Dim hFighterToReplace As HtmlElement = FindFighterHTMLContainer(lstPositions(lstPos).orders(order).Last, hDoc)
            'получаем родителя удаляемого бойца
            Dim hParent As HtmlElement = hFighter.Parent
            'удаляем бойца
            Dim msFighter As mshtml.IHTMLDOMNode = hFighter.DomElement
            msFighter.removeNode(True)
            If IsNothing(hFighterToReplace) = False Then
                'определяем не является ли боец, который встанет на замену, единственным в свойе ячейке
                Dim hRepParent As HtmlElement = hFighterToReplace.Parent
                Dim isOnly As Boolean = (hRepParent.Children.Count = 1)

                'перемещяем бойца для замены
                Dim msFighterToReplace As mshtml.IHTMLDOMNode = hFighterToReplace.DomElement
                msFighterToReplace.removeNode(True)
                hParent.AppendChild(hFighterToReplace)

                'если перемещяемый был единственным в своей ячейке - удаляем ячейку
                If isOnly Then
                    Dim msRepParent As mshtml.IHTMLDOMNode = hRepParent.DomElement
                    msRepParent.removeNode(True)
                End If
            End If
        End If

        'обновляем структуру с расположением бойцов
        If PrepareFightersPosition() = "#Error" Then Return "#Error"
        Return ""
    End Function

    ''' <summary>
    ''' Функция возвращает html-элемент IMG с изображением указанного бойца
    ''' </summary>
    ''' <param name="hFighterDiv">Id html-контейнера бойца</param>
    ''' <param name="fId">Id бойца</param>
    Public Function FindFighterHTMLImage(ByRef hFighterDiv As HtmlElement, ByVal fId As Integer, ByRef hEffect As HtmlElement) As HtmlElement
        'получаем html-элемент картинки
        Dim hParent As HtmlElement = hFighterDiv
        If hFighterDiv.Children(0).TagName = "TABLE" Then
            Dim VerticalLines As String = hFighterDiv.Children(0).GetAttribute("VerticalLines")
            If VerticalLines = "both" OrElse VerticalLines = "left" Then
                hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
            ElseIf VerticalLines = "right" Then
                hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
            End If
        End If

        Dim heroImg As HtmlElement = Nothing
        For i As Integer = 0 To hParent.Children.Count - 1
            If hParent.Children(i).TagName = "IMG" Then
                Dim imageType As String = hParent.Children(i).GetAttribute("imageType")
                If imageType = "hero" Then
                    heroImg = hParent.Children(i)
                ElseIf imageType = "effect" Then
                    hEffect = hParent.Children(i)
                End If
                'ElseIf hParent.Children(i).TagName = "DIV" Then
                '    Dim imageType As String = hParent.Children(i).GetAttribute("imageType")
                '    If imageType = "effect" Then
                '        hEffect = hParent.Children(i)
                '    End If
            End If
        Next

        Return heroImg
    End Function

    ''' <summary>
    ''' Печатает иконки способностей у бойца
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="excludeId">Id способности, которую отображать не надо (при удалении)</param>
    Public Function ChangeAbilitiesIcons(ByVal fId As Integer, Optional ByVal excludeId As Integer = -1) As String
        Dim classB As Integer = mScript.mainClassHash("B")
        Dim nameToSeek As String = "[Abilities]"

        'ищем свойство битвы PropertyToDisplayХ, которое содержит [Abilities]
        Dim i As Integer = 1, displayNum As Integer = 0, res As String = ""
        Do
            Dim pName As String = "PropertyToDisplay" & i.ToString
            If mScript.mainClass(classB).Properties.ContainsKey(pName) Then
                If ReadProperty(classB, pName, -1, -1, res, Nothing) = False Then Return "#Error"
                If UnWrapString(res) = nameToSeek Then
                    displayNum = i
                    Exit Do
                End If
            Else
                Exit Do
            End If
            i += 1
        Loop
        If displayNum = 0 Then Return "" 'свойство не отображается в виде линии - выход

        'свойство должно отображаться
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не загружен html-документ в главное окно.")
        Dim hFighterDiv As HtmlElement = FindFighterHTMLContainer(fId, hDoc)
        If IsNothing(hFighterDiv) Then Return ""
        Dim hParent As HtmlElement = hFighterDiv

        'получаем положение (стиль) линии
        Dim displayStyle As PropertyVisualizationStyleEnum
        If ReadPropertyInt(classB, "PropertyToDisplay" & displayNum.ToString & "Style", -1, -1, displayStyle, Nothing) = False Then Return "#Error"
        If displayStyle = PropertyVisualizationStyleEnum.VERTICAL_FRONT Then
            Dim fLeft As Boolean = CBool(hFighterDiv.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TD-TR-TBODY-TABLE
            If fLeft Then
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
            Else
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_LEFT
            End If
        ElseIf displayStyle = PropertyVisualizationStyleEnum.VERTICAL_BACK Then
            Dim fLeft As Boolean = CBool(hFighterDiv.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TD-TR-TBODY-TABLE
            If fLeft Then
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_LEFT
            Else
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
            End If
        End If

        'получаем родительский элемент, в котором находится линия
        If hFighterDiv.Children(0).TagName = "TABLE" Then
            Dim VerticalLines As String = hFighterDiv.Children(0).GetAttribute("VerticalLines")
            Select Case VerticalLines
                Case "both"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(2) 'TABLE-TBODY-TR-TD(2)
                    End Select
                Case "left"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                    End Select
                Case "right"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                    End Select
            End Select
        End If

        Dim hParamTotal As HtmlElement = Nothing
        For i = 0 To hParent.Children.Count - 1
            Dim ch As HtmlElement = hParent.Children(i)
            If ch.TagName = "DIV" Then
                Dim pName As String = ch.GetAttribute("PropertyName")
                If pName = nameToSeek Then
                    hParamTotal = ch
                    Exit For
                End If
            ElseIf ch.TagName = "TABLE" Then
                'перебираем все TR
                For j As Integer = 0 To ch.Children(0).Children.Count - 1
                    'перебираем все TD
                    For u As Integer = 0 To ch.Children(0).Children(j).Children.Count - 1
                        Dim ch2 As HtmlElement = ch.Children(0).Children(j).Children(u)
                        'перебираем всех детей внутри ячейки
                        For q As Integer = 0 To ch2.Children.Count - 1
                            If ch2.Children.Count = 0 Then Continue For
                            Dim pName As String = ch2.Children(0).GetAttribute("PropertyName")
                            If pName = nameToSeek Then
                                hParamTotal = ch2.Children(0)
                                Exit For
                            End If
                        Next q
                        If IsNothing(hParamTotal) = False Then Exit For
                    Next u
                    If IsNothing(hParamTotal) = False Then Exit For
                Next j
            End If
            If IsNothing(hParamTotal) = False Then Exit For
        Next i

        If IsNothing(hParamTotal) Then Return ""
        Return PrintAbilitiesIcons(hParamTotal, fId, displayStyle, excludeId)
    End Function

    ''' <summary>
    ''' Если свойство отображается в виде линии возле бойца, то изменяем ее визуальное представление при изменении значения свойства
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="paramName">Имя отображаемого свойства (без кавычек)</param>
    Public Function ChangeFighterParam(ByVal fId As Integer, ByVal paramName As String) As String
        Dim classB As Integer = mScript.mainClassHash("B")
        Dim nameToSeek As String = paramName
        If nameToSeek.EndsWith("Total") Then nameToSeek = paramName.Substring(0, paramName.Length - 5)

        'ищем свойство битвы PropertyToDisplayХ, которое описывает даное свойство
        Dim i As Integer = 1, displayNum As Integer = 0, res As String = ""
        Do
            Dim pName As String = "PropertyToDisplay" & i.ToString
            If mScript.mainClass(classB).Properties.ContainsKey(pName) Then
                If ReadProperty(classB, pName, -1, -1, res, Nothing) = False Then Return "#Error"
                If UnWrapString(res) = nameToSeek Then
                    displayNum = i
                    Exit Do
                End If
            Else
                Exit Do
            End If
            i += 1
        Loop
        If displayNum = 0 Then Return "" 'свойство не отображается в виде линии - выход
        Dim propValue As Double = 0, arrs() As String = {fId.ToString}
        If ReadFighterProperty(nameToSeek & "Total", fId, propValue, arrs) = False Then Return "#Error"
        If propValue <= 0 Then Return "" 'если максимум [propName]Total = 0, то свойство в виде линии не отображается

        'свойство должно отображаться
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не загружен html-документ в главное окно.")
        Dim hFighterDiv As HtmlElement = FindFighterHTMLContainer(fId, hDoc)
        If IsNothing(hFighterDiv) Then Return ""
        Dim hParent As HtmlElement = hFighterDiv

        'получаем положение (стиль) линии
        Dim displayStyle As PropertyVisualizationStyleEnum
        If ReadPropertyInt(classB, "PropertyToDisplay" & displayNum.ToString & "Style", -1, -1, displayStyle, Nothing) = False Then Return "#Error"
        If displayStyle = PropertyVisualizationStyleEnum.VERTICAL_FRONT Then
            Dim fLeft As Boolean = CBool(hFighterDiv.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TD-TR-TBODY-TABLE
            If fLeft Then
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
            Else
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_LEFT
            End If
        ElseIf displayStyle = PropertyVisualizationStyleEnum.VERTICAL_BACK Then
            Dim fLeft As Boolean = CBool(hFighterDiv.Parent.Parent.Parent.Parent.GetAttribute("fromLeft")) '..-TD-TR-TBODY-TABLE
            If fLeft Then
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_LEFT
            Else
                displayStyle = PropertyVisualizationStyleEnum.VERTICAL_RIGHT
            End If
        End If

        'получаем родительский элемент, в котором находится линия
        If hFighterDiv.Children(0).TagName = "TABLE" Then
            Dim VerticalLines As String = hFighterDiv.Children(0).GetAttribute("VerticalLines")
            Select Case VerticalLines
                Case "both"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(2) 'TABLE-TBODY-TR-TD(2)
                    End Select
                Case "left"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                        Case PropertyVisualizationStyleEnum.VERTICAL_LEFT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                    End Select
                Case "right"
                    Select Case displayStyle
                        Case PropertyVisualizationStyleEnum.HORIZONTAL_TOP, PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD(0)
                        Case PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                            hParent = hFighterDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
                    End Select
            End Select
        End If

        Dim hParamTotal As HtmlElement = Nothing
        For i = 0 To hParent.Children.Count - 1
            Dim ch As HtmlElement = hParent.Children(i)
            If ch.TagName = "DIV" Then
                Dim pName As String = ch.GetAttribute("PropertyName")
                If pName = nameToSeek Then
                    hParamTotal = ch
                    Exit For
                End If
            ElseIf ch.TagName = "TABLE" Then
                'перебираем все TR
                For j As Integer = 0 To ch.Children(0).Children.Count - 1
                    'перебираем все TD
                    For u As Integer = 0 To ch.Children(0).Children(j).Children.Count - 1
                        Dim ch2 As HtmlElement = ch.Children(0).Children(j).Children(u)
                        'перебираем всех детей внутри ячейки
                        For q As Integer = 0 To ch2.Children.Count - 1
                            If ch2.Children.Count = 0 Then Continue For
                            Dim pName As String = ch2.Children(0).GetAttribute("PropertyName")
                            If pName = nameToSeek Then
                                hParamTotal = ch2.Children(0)
                                Exit For
                            End If
                        Next q
                        If IsNothing(hParamTotal) = False Then Exit For
                    Next u
                    If IsNothing(hParamTotal) = False Then Exit For
                Next j
            End If
            If IsNothing(hParamTotal) = False Then Exit For
        Next i

        If IsNothing(hParamTotal) Then Return ""

        'параметр найден
        'hParamTotal(hParam(hParamInner))
        Dim hParam As HtmlElement = hParamTotal.Children(0)
        If IsNothing(hParam) Then Return ""
        Dim hParamInner As HtmlElement = hParam.Children(0)
        If IsNothing(hParamInner) Then Return ""

        Dim value As Double = 0, valueTotal As Double = 0
        If ReadFighterPropertyDbl(nameToSeek, fId, value, arrs) = False Then Return "#Error"
        If ReadFighterPropertyDbl(nameToSeek & "Total", fId, valueTotal, arrs) = False Then Return "#Error"
        Dim textBeforeNew As String = "", textAfterNew As String = ""
        If mScript.mainClass(classB).Properties.ContainsKey("PropertyToDisplayBefore" & displayNum.ToString) Then
            If ReadProperty(classB, "PropertyToDisplayBefore" & displayNum.ToString, -1, -1, textBeforeNew, arrs) = False Then Return "#Error"
            textBeforeNew = UnWrapString(textBeforeNew)
        End If
        If mScript.mainClass(classB).Properties.ContainsKey("PropertyToDisplayAfter" & displayNum.ToString) Then
            If ReadProperty(classB, "PropertyToDisplayAfter" & displayNum.ToString, -1, -1, textAfterNew, arrs) = False Then Return "#Error"
            textAfterNew = UnWrapString(textAfterNew)
        End If

        'получаем разницу в % между текущим и максимальным значением
        If valueTotal <= 0 Then Return ""
        Dim valueDifference As Integer, lineWidth As String = ""
        valueDifference = Math.Round(100 * value / valueTotal, 0)
        If valueDifference > 100 Then
            valueDifference = 100
        ElseIf valueDifference < 0 Then
            valueDifference = 0
        End If

        Select Case displayStyle
            Case PropertyVisualizationStyleEnum.VERTICAL_LEFT, PropertyVisualizationStyleEnum.VERTICAL_RIGHT
                Dim textOld As String, attrText As String = "", hTDText As HtmlElement = Nothing
                Dim blnFoundBefore As Boolean = False, blnFoundAfter As Boolean = False 'существует и надо ли печатать текст перед/после линии
                If String.IsNullOrEmpty(textBeforeNew) Then blnFoundBefore = True
                If String.IsNullOrEmpty(textAfterNew) Then blnFoundAfter = True

                'менеям text before (если найдем ячейку)
                If hParamTotal.Parent.TagName = "TD" Then
                    hTDText = hParamTotal.Parent.Parent.Parent.Children(0).Children(0) '-..TD-TR(x)-TBODY-TR(0)-TD(0)
                    attrText = hTDText.GetAttribute("text")
                    If attrText = "before" Then
                        'ячейка для текста найдена
                        textOld = hTDText.InnerHtml
                        If textBeforeNew <> textOld Then hTDText.InnerHtml = textBeforeNew
                        blnFoundBefore = True
                    End If
                End If

                'меняем textAfter (если найдем ячейку)
                hTDText = hParamTotal.Parent.Parent.Parent '-..TD-TR(x)-TBODY
                If hTDText.Children.Count > 1 Then
                    hTDText = hTDText.Children(hTDText.Children.Count - 1).Children(0) '-..TR(2)-TD(0)
                    attrText = hTDText.GetAttribute("text")
                    If attrText = "after" Then
                        'ячейка для текста найдена
                        textOld = hTDText.InnerHtml
                        If textAfterNew <> textOld Then hTDText.InnerHtml = textAfterNew
                        blnFoundAfter = True
                    End If
                End If

                If blnFoundAfter = False OrElse blnFoundBefore = False Then
                    'таблицы с текстом до/после линии не существует - перерисовываем бойца
                    Return PutFighterToBattleField(fId, hFighterDiv.Parent, hFighterDiv)
                End If

                'меняем класс graphProp_ на graphPropExtra_ или наоборот при необходимости
                Dim strClass As String = hParamInner.GetAttribute("ClassName")
                If value <= valueTotal Then
                    'должен быть обычный класс
                    If strClass.StartsWith("graphPropExtra_") Then
                        'заменяем класс экстра на обычный
                        hParamInner.SetAttribute("className", Replace(strClass, "graphPropExtra_", "graphProp_", 1, 1, CompareMethod.Text))
                    End If
                Else
                    'должен быть класс экстра
                    If strClass.StartsWith("graphProp_") Then
                        'заменяем обычный класс на класс экстра
                        hParamInner.SetAttribute("className", Replace(strClass, "graphProp_", "graphPropExtra_", 1, 1, CompareMethod.Text))
                    End If
                End If

                'меняем ширину линии в зависимости от значения valueDifference
                lineWidth = (100 - valueDifference).ToString & "%"
                hParamInner.Style &= "position:relative;top:" & lineWidth & ";"
            Case Else
                Dim textOld As String, attrText As String = "", hTDText As HtmlElement = Nothing
                Dim blnFoundBefore As Boolean = False, blnFoundAfter As Boolean = False 'существует и надо ли печатать текст перед/после линии
                If String.IsNullOrEmpty(textBeforeNew) Then blnFoundBefore = True
                If String.IsNullOrEmpty(textAfterNew) Then blnFoundAfter = True
                Dim TDBeforeExists As Boolean = False, TDAfterExists As Boolean = False 'существует ли ячейка таблицы перед/после линии

                'менеям text before (если найдем ячейку)
                If hParamTotal.Parent.TagName = "TD" Then
                    hTDText = hParamTotal.Parent.Parent.Children(0) '-..TD-TR(0)-TD(0)
                    attrText = hTDText.GetAttribute("text")
                    If attrText = "before" Then
                        'ячейка для текста найдена
                        textOld = hTDText.InnerHtml
                        If textBeforeNew <> textOld Then hTDText.InnerHtml = textBeforeNew
                        blnFoundBefore = True
                        TDBeforeExists = True
                    End If
                End If

                'меняем textAfter (если найдем ячейку)
                hTDText = hParamTotal.Parent.Parent '-..TD-TR(0)
                If hTDText.Children.Count > 1 Then
                    hTDText = hTDText.Children(hTDText.Children.Count - 1) '-..-TD(2)
                    attrText = hTDText.GetAttribute("text")
                    If attrText = "after" Then
                        'ячейка для текста найдена
                        textOld = hTDText.InnerHtml
                        If textAfterNew <> textOld Then hTDText.InnerHtml = textAfterNew
                        blnFoundAfter = True
                        TDAfterExists = True
                    End If
                End If

                If TDAfterExists = False AndAlso TDBeforeExists = False Then
                    If (hParamTotal.Parent.TagName = "DIV" AndAlso blnFoundAfter AndAlso blnFoundBefore) = False Then
                        'таблицы с текстом до/после линии не существует и она нужна - перерисовываем бойца
                        Return PutFighterToBattleField(fId, hFighterDiv.Parent, hFighterDiv)
                    End If
                End If

                'меняем класс graphProp_ на graphPropExtra_ или наоборот при необходимости
                Dim strClass As String = hParamInner.GetAttribute("ClassName")
                If value <= valueTotal Then
                    'должен быть обычный класс
                    If strClass.StartsWith("graphPropExtra_") Then
                        'заменяем класс экстра на обычный
                        hParamInner.SetAttribute("className", Replace(strClass, "graphPropExtra_", "graphProp_", 1, 1, CompareMethod.Text))
                    End If
                Else
                    'должен быть класс экстра
                    If strClass.StartsWith("graphProp_") Then
                        'заменяем обычный класс на класс экстра
                        hParamInner.SetAttribute("className", Replace(strClass, "graphProp_", "graphPropExtra_", 1, 1, CompareMethod.Text))
                    End If
                End If

                'меняем ширину линии в зависимости от значения valueDifference
                lineWidth = valueDifference.ToString & "%"
                hParam.Style &= "width:" & lineWidth & ";height:100%;"

                If blnFoundBefore = False OrElse blnFoundAfter = False Then
                    'не существует нужной ячейки (а сама таблица существует)
                    Dim TR As HtmlElement = Nothing, TDline As HtmlElement
                    'получаем TR
                    TR = hParamTotal.Parent.Parent '..-TD-TR
                    TDline = hParamTotal.Parent 'ячейка, в которой находится линия

                    If blnFoundBefore = False Then
                        'создаем ячейку с Before
                        Dim TD As HtmlElement = hDoc.CreateElement("TD")
                        If strClass.StartsWith("graphProp_") Then
                            TD.SetAttribute("ClassName", Replace(strClass, "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                        Else
                            TD.SetAttribute("ClassName", Replace(strClass, "graphPropExtra_", "graphPropText_", 1, 1, CompareMethod.Text))
                        End If
                        TD.InnerHtml = textBeforeNew
                        TD.SetAttribute("text", "before")
                        TDline.InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, TD)
                    End If

                    If blnFoundAfter = False Then
                        'создаем ячейку с After
                        Dim TD As HtmlElement = hDoc.CreateElement("TD")
                        If strClass.StartsWith("graphProp_") Then
                            TD.SetAttribute("ClassName", Replace(strClass, "graphProp_", "graphPropText_", 1, 1, CompareMethod.Text))
                        Else
                            TD.SetAttribute("ClassName", Replace(strClass, "graphPropExtra_", "graphPropText_", 1, 1, CompareMethod.Text))
                        End If
                        TD.InnerHtml = textAfterNew
                        TD.SetAttribute("text", "after")
                        TR.AppendChild(TD)
                    End If
                End If
        End Select

        Return ""
    End Function

    ''' <summary>
    ''' Печатает иконки способностей Ab.BattleIcon, которыми обладает указанный персонаж
    ''' </summary>
    ''' <param name="hDest">html-элемент, в котором надо разместить иконки</param>
    ''' <param name="hId">Id героя</param>
    ''' <param name="pStyle">стиль контейнера для отображения (вертикальный или горизонтальный)</param>
    ''' <param name="excludeId">Id способности, которую отображать не надо (при удалении)</param>
    Private Function PrintAbilitiesIcons(ByRef hDest As HtmlElement, ByVal hId As Integer, ByVal pStyle As PropertyVisualizationStyleEnum, Optional ByVal excludeId As Integer = -1) As String
        hDest.InnerHtml = ""
        Dim abSetId As Integer = Fighters(hId).abilitySetId
        If abSetId < 0 Then Return ""

        Dim classAb As Integer = mScript.mainClassHash("Ab")
        Dim ch As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classAb).ChildProperties(abSetId)
        If IsNothing(ch("Name").ThirdLevelProperties) OrElse ch("Name").ThirdLevelProperties.Count = 0 OrElse (ch("Name").ThirdLevelProperties.Count = -1 AndAlso excludeId > -1) Then Return ""

        Dim arrs() As String = {abSetId.ToString, ""}
        For abId As Integer = 0 To ch("Name").ThirdLevelProperties.Count - 1
            'перебираем все способности персонажа
            If abId = excludeId Then Continue For
            arrs(1) = abId.ToString

            'получаем BattleIcon. Если не указана, то пропускаем способность
            Dim BattleIcon As String = ""
            If ReadProperty(classAb, "BattleIcon", abSetId, abId, BattleIcon, arrs) = False Then Return "#Error"
            BattleIcon = UnWrapString(BattleIcon)
            If String.IsNullOrEmpty(BattleIcon) OrElse My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, BattleIcon)) = False Then Continue For
            BattleIcon = BattleIcon.Replace("\"c, "/"c)

            'получаем Visible. Если False, то пропускаем способность
            Dim visible As Boolean = True
            If ReadPropertyBool(classAb, "Visible", abSetId, abId, visible, arrs) = False Then Return "#Error"
            If Not visible Then Continue For

            'получаем Enabled. Если False, то пропускаем способность
            Dim enabled As Boolean = True
            If ReadPropertyBool(classAb, "Enabled", abSetId, abId, enabled, arrs) = False Then Return "#Error"
            If Not enabled Then Continue For

            'получаем подпись и описание
            Dim capt As String = "", desc As String = ""
            If ReadProperty(classAb, "Caption", abSetId, abId, capt, arrs) = False Then Return "#Error"
            capt = UnWrapString(capt)
            If ReadProperty(classAb, "Description", abSetId, abId, desc, arrs) = False Then Return "#Error"
            desc = UnWrapString(desc)
            If String.IsNullOrEmpty(capt) Then
                capt = desc
            ElseIf String.IsNullOrEmpty(desc) = False Then
                capt &= ". " & desc
            End If

            'создаем иконку
            Dim hImg As HtmlElement = hDest.Document.CreateElement("IMG")
            If pStyle = PropertyVisualizationStyleEnum.HORIZONTAL_BOTTOM OrElse pStyle = PropertyVisualizationStyleEnum.HORIZONTAL_TOP Then
                'горизонтальная линия
                hImg.Style = "display:inline-block"
            Else
                'вертикальная линия
                hImg.Style = "display:block"
            End If
            hImg.SetAttribute("src", BattleIcon)
            hImg.SetAttribute("Title", capt)
            hDest.AppendChild(hImg)
        Next abId

        Return ""
    End Function

    ''' <summary>
    ''' Подготавливает список бойцов учавствующих в бою, разделяя их по армиям и рангам
    ''' </summary>
    Public Function PrepareFightersPosition() As String
        lstPositions.Clear()
        Dim classArmy As Integer = mScript.mainClassHash("Army")

        For fId As Integer = 0 To Fighters.Count - 1
            'перебираем всех бойцов
            Dim arrs() As String = {fId.ToString}
            Dim armyId As Integer = Fighters(fId).armyId
            Dim fromLeft As Boolean = True
            Dim pos As Integer = -1 'Id в lstPositions
            'сразу получаем ранг текущего бойца
            Dim range As Integer = 0, life As Double = 0, runAway As Boolean = False
            If ReadFighterPropertyInt("Range", fId, range, arrs) = False Then Return "#Error"
            If ReadFighterPropertyDbl("Life", fId, life, arrs) = False Then Return "#Error"
            If ReadFighterPropertyBool("RunAway", fId, runAway, arrs) = False Then Return "#Error"
            If life <= 0 OrElse runAway Then Continue For

            If armyId < 0 Then
                'боец без армии
                'получаем свойство Friendliness, чтобы установить fromLeft
                Dim fr As Integer = 0
                If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return "#Error"
                If fr = 2 Then
                    'призванный - получаем Friendliness хозяина
                    Dim hostId As Integer = GetSummonOwner(fId)
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return "#Error"
                End If

                fromLeft = Not (fr = 0)

                'если запись в lstPositions (без армии и с тем же fromLeft) существует, то получаем ее
                For i As Integer = 0 To lstPositions.Count - 1
                    If lstPositions(i).armyId = -1 AndAlso lstPositions(i).fromLeft = fromLeft Then
                        pos = i
                        Exit For
                    End If
                Next i
            Else
                'персонаж в армии
                'если запись в lstPositions с такой же армией уже существует - получаем ее
                For i As Integer = 0 To lstPositions.Count - 1
                    If lstPositions(i).armyId = armyId Then
                        pos = i
                        Exit For
                    End If
                Next i
            End If

            'уже получено Id армии и fromLeft, а также pos (Id в lstPositions)
            Dim fPos As FightersPosition
            Dim orders As SortedList(Of Integer, List(Of Integer))
            If pos = -1 Then
                'такой армии еще не было
                'получаем Position армии (1 - слева, 2 - справа, 0 - авто на основании Friendliness первого бойца)
                If armyId > -1 Then
                    Dim armyPos As Integer = 0
                    If ReadPropertyInt(classArmy, "Position", armyId, -1, armyPos, {armyId.ToString}) = False Then Return "#Error"

                    If armyPos = 1 Then
                        fromLeft = True
                    ElseIf armyPos = 2 Then
                        fromLeft = False
                    Else
                        'авто на основании Friendliness первого бойца)
                        Dim fr As Integer = 0
                        If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return "#Error"
                        If fr = 2 Then
                            'призванный - получаем Friendliness хозяина
                            Dim hostId As Integer = GetSummonOwner(fId)
                            If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return "#Error"
                        End If

                        fromLeft = Not (fr = 0)
                    End If
                End If

                'создаем новую запись в lstPositions
                pos = lstPositions.Count
                fPos = New FightersPosition With {.armyId = armyId, .fromLeft = fromLeft}
                orders = New SortedList(Of Integer, List(Of Integer)) 'ключ - ранг, значение - список бойцов с данным рангом
                Dim lst As New List(Of Integer) 'новый список бойцов
                lst.Add(fId)
                orders.Add(range, lst)
                fPos.orders = orders
                lstPositions.Add(fPos)
            Else
                'такая армия уже есть
                fPos = lstPositions(pos)
                orders = fPos.orders 'ключ - ранг, значение - список бойцов с данным рангом

                Dim lst As List(Of Integer)
                If orders.ContainsKey(range) Then
                    'уже есть бойцы с таким же рангом
                    lst = orders(range)
                    lst.Add(fId)
                Else
                    'бойцов с аналогичным рангом пока нет
                    lst = New List(Of Integer)
                    lst.Add(fId)
                    orders.Add(range, lst)
                End If
            End If
        Next

        Return ""
    End Function

    ''' <summary>
    ''' Получает ссылку на html-элемент армии, ее названия и герба
    ''' </summary>
    ''' <param name="armyId">Id армии</param>
    ''' <param name="hCaption">Для получения html-элемента названия</param>
    ''' <param name="hCoatOfArms">Для получения html-элемента герба</param>
    Public Function FindArmyHTMLContainer(ByVal armyId As Integer, ByRef hCaption As HtmlElement, ByRef hCoatOfArms As HtmlElement) As HtmlElement
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then
            mScript.LAST_ERROR = "Не загружен html-документ в главное окно."
            Return Nothing
        End If

        For z As Integer = 1 To 2
            'ищем в контейнерах для левоых и правых армий
            Dim ArmiesConvas As HtmlElement = IIf(z = 1, hDoc.GetElementById("LeftArmies"), hDoc.GetElementById("RightArmies"))
            If IsNothing(ArmiesConvas) Then Return Nothing

            For i As Integer = 0 To ArmiesConvas.Children.Count - 1
                'перебираем все армии в контейнере
                Dim hArmy As HtmlElement = ArmiesConvas.Children(i)
                Dim army As String = hArmy.GetAttribute("armyId")
                If String.IsNullOrEmpty(army) = False AndAlso Val(army) = armyId Then
                    'армия найдена
                    'получаем ссылки на html-элементы названия и герба армии
                    Dim hEL As HtmlElement = hArmy.Children(0)
                    If hEL.GetAttribute("ClassName") = "ArmyCaption" Then
                        hCaption = hEL
                        hEL = hArmy.Children(1)
                        If hEL.GetAttribute("ClassName") = "ArmyCoatOfArms" Then hCoatOfArms = hEL
                    ElseIf hEL.GetAttribute("ClassName") = "ArmyCoatOfArms" Then
                        hCoatOfArms = hEL
                    End If
                    Return hArmy
                End If
            Next i
        Next z
        Return Nothing
    End Function

    ''' <summary>
    ''' Возвращает html-элемент (DIV), являющийся контейнером указанного героя
    ''' </summary>
    ''' <param name="fId">Id персонажа</param>
    ''' <param name="hDoc">html-документ</param>
    Public Function FindFighterHTMLContainer(ByVal fId As Integer, ByRef hDoc As HtmlDocument) As HtmlElement
        If IsNothing(hDoc) Then hDoc = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then
            mScript.LAST_ERROR = "Не загружен html-документ в главное окно."
            Return Nothing
        End If

        For z As Integer = 1 To 2
            'ищем в контейнерах для левоых и правых армий
            Dim ArmiesConvas As HtmlElement = IIf(z = 1, hDoc.GetElementById("LeftArmies"), hDoc.GetElementById("RightArmies"))
            If IsNothing(ArmiesConvas) Then Return Nothing

            For i As Integer = 0 To ArmiesConvas.Children.Count - 1
                'перебираем все армии в контейнере
                Dim hArmy As HtmlElement = ArmiesConvas.Children(i)
                Dim army As String = hArmy.GetAttribute("armyId")
                If String.IsNullOrEmpty(army) OrElse Val(army) <> Fighters(fId).armyId Then Continue For 'если армия персонажа другая - пропускаем армию

                'контейнер армии найден. В нем TABLE-TBODY-TR-TD(x). А в начале может быть подпись и герб
                Dim hTable As HtmlElement = hArmy.Children(hArmy.Children.Count - 1)
                If hTable.TagName <> "TABLE" Then Continue For

                Dim TR As HtmlElement = hTable.Children(0).Children(0)
                For j As Integer = 0 To TR.Children.Count - 1
                    'перебираем все ячейки таблицы, каждая из которых содержит бойцов определенного ранга
                    Dim TD As HtmlElement = TR.Children(j)
                    If TD.Children.Count = 0 Then Continue For
                    For q As Integer = 0 To TD.Children.Count - 1
                        'перебираем бойцов внутри ячейки
                        Dim hero As String = TD.Children(q).GetAttribute("heroId")
                        If String.IsNullOrEmpty(hero) Then Continue For
                        If Val(hero) = fId Then Return TD.Children(q)
                    Next q
                Next j
            Next i
        Next z
        Return Nothing
    End Function

    ''' <summary>
    ''' Получает бойца, который должен атаковать следующим
    ''' </summary>
    ''' <param name="fighterId">Id текущего бойца или -1 если в этом раунде еще никто не ходил</param>
    ''' <param name="AddTurnToPrevFighter">Указывать что текущий боец походил или нет</param>
    ''' <returns>-2 при ошибке, -1 если все в этом раунде походили, Id следующего бойца</returns>
    Public Function GetNextFighter(ByVal fighterId As Integer, Optional AddTurnToPrevFighter As Boolean = True) As Integer
        'Dim classB As Integer = mScript.mainClassHash("B")
        Dim maxTurns As Integer = 1
        'If ReadPropertyInt(classB, "AttacksPerTurn", -1, -1, maxTurns, Nothing) = False Then Return -2

        If AddTurnToPrevFighter AndAlso fighterId > -1 Then
            'увеличиваем на 1 количество ходов
            Dim prop As MatewScript.ChildPropertiesInfoType = Fighters(fighterId).heroProps("AttacksInThisTurn")
            Dim prevValue As Integer = Val(prop.Value)
            prop.Value = (prevValue + 1).ToString
        End If

        'получаем список бойцов, которые могут походить в этом раунде
        Dim lstFighters As New List(Of CandidatesToBeNext) 'для Id бойцов, которые могут походить
        Dim attacks As Integer = 0, speed As Double = 0, runAway As Boolean = False, life As Double = 0
        Dim minTurns As Integer = Integer.MaxValue
        For fId As Integer = 0 To Fighters.Count - 1
            Dim arrs() As String = {fId.ToString}
            If ReadFighterPropertyInt("AttacksPerTurn", fId, maxTurns, arrs) = False Then Return -2
            If ReadFighterPropertyInt("AttacksInThisTurn", fId, attacks, arrs) = False Then Return -2
            If attacks >= maxTurns Then Continue For

            If ReadFighterPropertyBool("RunAway", fId, runAway, arrs) = False Then Return -2
            If runAway Then Continue For

            If ReadFighterPropertyDbl("Life", fId, life, arrs) = False Then Return -2
            If life <= 0 Then Continue For

            If ReadFighterPropertyDbl("Speed", fId, speed, arrs) = False Then Return -2
            If attacks < minTurns Then
                minTurns = attacks
            Else
                'этот боец походить может, но он походил большее количество раз, чем один из других (он точно не следующий). Нет смысла его добавлять в lstFighters
                Continue For
            End If
            lstFighters.Add(New CandidatesToBeNext With {.fighterId = fId, .speed = speed, .turns = attacks})
        Next fId

        If lstFighters.Count = 0 Then Return -1 'все походили

        'получаем среди бойцов с наименьшим значением turns самого быстрого
        Dim maxSpeed As Double = Double.MinValue
        Dim nextFighter As Integer = -1
        For i As Integer = 0 To lstFighters.Count - 1
            If lstFighters(i).turns > minTurns Then Continue For
            If lstFighters(i).speed > maxSpeed Then
                maxSpeed = lstFighters(i).speed
                nextFighter = lstFighters(i).fighterId
            End If
        Next i

        Return nextFighter
    End Function

    ''' <summary>
    ''' Возвращает количество бойцов
    ''' </summary>
    ''' <param name="aliveOnly">считать только живых (и не сбежавших)</param>
    ''' <param name="fType">тип бойцов: 0 - все, 1 - друзья, 2 - враги, 3 - все призванные, 4 - призванные ГГ, 5 - призванные друзьями, 6 - призванные врагами, 7 - главные герои</param>
    ''' <returns>-1 при ошибке</returns>
    Public Function FightersCount(ByVal aliveOnly As Boolean, ByVal fType As Integer) As Integer
        'ByVal fAlive As Boolean, ByVal fType As Long
        '0 - Alive
        '1 - Type:
        '0 - All
        '1 - Friends
        '2 - Enemies
        '3 - All Summoned
        '4 - Summoned By You
        '5 - Summoned By Friends
        '6 - Summoned By Enemies
        '7 - ГГ

        If Fighters.Count = 0 Then Return 0
        If aliveOnly = False AndAlso fType = 0 Then Return Fighters.Count

        Dim cnt As Integer = 0, fLife As Double = 0, runAway As Boolean = False, fEnemy As IsEnemyEnum
        For fId As Integer = 0 To Fighters.Count - 1
            Dim arrs() As String = {fId.ToString}
            If aliveOnly Then
                'выбираем только живых
                If ReadFighterPropertyDbl("Life", fId, fLife, arrs) = False Then Return -1
                If ReadFighterPropertyBool("RunAway", fId, runAway, arrs) = False Then Return -1
                If fLife <= 0 OrElse runAway Then Continue For
            End If

            'выбираем тип
            Select Case fType
                Case 0 'All
                    cnt += 1
                Case 1 'Friends
                    fEnemy = IsEnemy(-1, fId)
                    If fEnemy = IsEnemyEnum.Friendly Then
                        cnt += 1
                    ElseIf fEnemy = IsEnemyEnum.CheckError Then
                        Return -1
                    End If
                Case 2 'Enemies
                    fEnemy = IsEnemy(-1, fId)
                    If fEnemy = IsEnemyEnum.Enemy Then
                        cnt += 1
                    ElseIf fEnemy = IsEnemyEnum.CheckError Then
                        Return -1
                    End If
                Case 3 'All Summoned
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return -1
                    If fr = 2 Then cnt += 1
                Case 4 'Summoned By You
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return -1
                    If fr = 2 Then
                        Dim ownerId As Integer = GetSummonOwner(fId)
                        If ownerId > -1 Then
                            If ReadFighterPropertyInt("Friendliness", ownerId, fr, arrs) = False Then Return -1
                            If fr = -1 Then cnt += 1
                        End If
                    End If
                Case 5 'Summoned By Friends
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return -1
                    If fr = 2 Then
                        Dim ownerId As Integer = GetSummonOwner(fId)
                        If ownerId > -1 Then
                            If ReadFighterPropertyInt("Friendliness", ownerId, fr, arrs) = False Then Return -1
                            If fr = 1 Then cnt += 1
                        End If
                    End If
                Case 6 'Summoned By Enemies
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return -1
                    If fr = 2 Then
                        Dim ownerId As Integer = GetSummonOwner(fId)
                        If ownerId > -1 Then
                            If ReadFighterPropertyInt("Friendliness", ownerId, fr, arrs) = False Then Return -1
                            If fr = 0 Then cnt += 1
                        End If
                    End If
                Case 7 'main heroes
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return -1
                    If fr = -1 Then cnt += 1
            End Select
        Next fId

        Return cnt
    End Function

    Public Function NewTurn() As String
        Dim classB As Integer = mScript.mainClassHash("B"), turnsCount As Integer = 0, eventId As Integer
        If ReadPropertyInt(classB, "TurnsCount", -1, -1, turnsCount, Nothing) = False Then Return "#Error"

        For fId As Integer = 0 To Fighters.Count - 1
            Fighters(fId).heroProps("AttacksInThisTurn").Value = "0"
        Next fId

        If turnsCount < 1 Then
            'Это первый раунд боя

            'проверка списка воспроизведения
            audioListBeforeBattle = GVARS.G_CURLIST
            Dim battleList As String = "", battleListId As Integer = -1
            Dim classMed As Integer = mScript.mainClassHash("Med")
            If ReadProperty(classB, "PlayList", -1, -1, battleList, Nothing) = False Then Return "#Error"
            battleListId = GetSecondChildIdByName(battleList, mScript.mainClass(classMed).ChildProperties)
            If battleListId >= 0 Then
                'меняем список воспроизведения
                If AudioStop() = False Then Return "#Error"
                GVARS.G_CURAUDIO = -1
                GVARS.G_CURLIST = battleListId
                If AudioPlayFromList(-1, Nothing) = "#Error" Then Return "#Error"
                timAudio.Enabled = True
            End If

            'запуск события BattleOnStart
            eventId = mScript.mainClass(classB).Properties("BattleOnStart").eventId
            If eventId > 0 Then
                If mScript.eventRouter.RunEvent(eventId, Nothing, "BattleOnStart", False) = "#Error" Then Return "#Error"
            End If
        Else
            'get starting hero (если раунд боя первый, то первый атакующий боец оперделяется в Battle.Begin)
            GVARS.G_CURFIGHTER = -1
            'Dim startingHero As String = "", startingHeroId As Integer = -1
            'If ReadProperty(classB, "StartingHero", -1, -1, startingHero, Nothing) = False Then Return "#Error"
            'startingHeroId = GetFighterByName(startingHero)
            'If startingHeroId < 0 Then
            '    GVARS.G_CURFIGHTER = GetNextFighter(-1)
            'Else
            '    GVARS.G_CURFIGHTER = startingHeroId
            'End If
        End If

        'проверка на окончание боя
        Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
        If bFinished <> BattleFinishEnum.CUSTOM Then
            If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
            Return StopBattle(bFinished)
        End If

        'увеличиваем turnsCount
        turnsCount += 1
        If PropertiesRouter(classB, "TurnsCount", Nothing, Nothing, PropertiesOperationEnum.PROPERTY_SET, turnsCount.ToString) = "#Error" Then Return "#Error"

        'событие BattleOnNewTurn
        eventId = mScript.mainClass(classB).Properties("BattleOnNewTurn").eventId
        If eventId > 0 Then
            If mScript.eventRouter.RunEvent(eventId, {turnsCount.ToString}, "BattleOnNewTurn", False) = "#Error" Then Return "#Error"
        End If

        'AbilityBattleTurnEvent
        Dim fLife As Double = 0
        Dim classAb As Integer = mScript.mainClassHash("Ab")
        Dim abEventIdGlobal As Integer, abEventIdHero As Integer
        abEventIdGlobal = mScript.mainClass(classAb).Properties("AbilityBattleTurnEvent").eventId
        For fId As Integer = 0 To Fighters.Count - 1
            'перебираем всех живых героев, имеющих наборы способностей
            'получаем набор способностей героя
            Dim abSetId As Integer = Fighters(fId).abilitySetId
            If abSetId < 0 Then Continue For
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityBattleTurnEvent")
            If IsNothing(ch.ThirdLevelEventId) OrElse ch.ThirdLevelEventId.Count = 0 Then Continue For

            'живой ли боец
            Dim arrs() As String = {fId.ToString}
            If ReadFighterPropertyDbl("Life", fId, fLife, arrs) = False Then Return "#Error"
            If fLife <= 0 Then Continue For

            abEventIdHero = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityBattleTurnEvent").eventId
            'перебираем все способности данного героя
            Dim abArrs() As String = {abSetId.ToString, "", fId.ToString, turnsCount.ToString}, res As String
            For abId As Integer = ch.ThirdLevelEventId.Count - 1 To 0 Step -1
                Dim arrProp() As String = {abArrs(0), abId.ToString}, enabled As Boolean = True, turnsActive As Integer = -1, abTurnsCount As Integer = 0
                'недоступна - пропускаем
                If ReadPropertyBool(classAb, "Enabled", abSetId, abId, enabled, arrProp) = False Then Return "#Error"
                If Not enabled Then Continue For

                'увеличиваем TurnsCount способности
                If ReadPropertyInt(classAb, "TurnsCount", abSetId, abId, abTurnsCount, arrProp) = False Then Return "#Error"
                abTurnsCount += 1
                If PropertiesRouter(classAb, "TurnsCount", arrProp, Nothing, PropertiesOperationEnum.PROPERTY_SET, abTurnsCount.ToString) = "#Error" Then Return "#Error"

                If ReadPropertyInt(classAb, "TurnsActive", abSetId, abId, turnsActive, arrProp) = False Then Return "#Error"
                If turnsActive > -1 Then
                    'способность ограничена по раундам боя
                    If abTurnsCount > turnsActive Then
                        'время действия способности вышло - снимаем способность
                        If FunctionRouter(classAb, "Remove", arrProp, Nothing) = "#Error" Then Return "#Error"
                        Continue For
                    End If
                End If

                res = ""
                abArrs(1) = abId.ToString
                'событие AbilityBattleTurnEvent
                If abEventIdGlobal > 0 Then
                    'глобальное
                    res = mScript.eventRouter.RunEvent(abEventIdGlobal, abArrs, "AbilityBattleTurnEvent", False)
                    If res = "False" Then
                        Continue For
                    ElseIf res = "#Error" Then
                        Return res
                    End If
                End If

                'набора способностей
                If abEventIdHero > 0 Then
                    res = mScript.eventRouter.RunEvent(abEventIdHero, abArrs, "AbilityBattleTurnEvent", False)
                    If res = "False" Then
                        Continue For
                    ElseIf res = "#Error" Then
                        Return res
                    End If
                End If

                'способности
                eventId = ch.ThirdLevelEventId(abId)
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, abArrs, "AbilityBattleTurnEvent", False)
                    If res = "#Error" Then Return res
                End If

                If res <> "False" Then
                    Dim snd As String = ""
                    If ReadProperty(classAb, "SoundOnActivate", abSetId, abId, snd, arrProp) = False Then Return "#Error"
                    snd = UnWrapString(snd)
                    If String.IsNullOrEmpty(snd) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, snd)) Then
                        'проигрываем звук
                        Say(snd, 100, False)
                    End If
                End If
            Next abId
        Next fId

        'проверка на окончание боя
        bFinished = CheckForBattleFinished()
        If bFinished <> BattleFinishEnum.CUSTOM Then
            If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
            Return StopBattle(bFinished)
        End If

        Return NextTurn(IIf(turnsCount <= 1, True, False), True)
    End Function

    Public Function NextTurn(Optional blnFirstTime As Boolean = False, Optional blnNewTurn As Boolean = False) As String
        'После нажатия "Дальше"
        Dim classB As Integer = mScript.mainClassHash("B")

        Dim prevFighter As Integer = -1, prevVictim As Integer = GVARS.G_CURVICTIM
        GVARS.G_CURVICTIM = -1
        GVARS.G_CANUSEMAGIC = True
        GVARS.G_STRIKETYPE = -1

        'получаем предыдущего бойца
        If blnFirstTime = False Then
            prevFighter = GVARS.G_CURFIGHTER
        End If
        'получаем нового бойца
        If blnNewTurn Then
            'первый ход - получаем StartingHero
            Dim sHero As String = ""
            If ReadProperty(classB, "StartingHero", -1, -1, sHero, Nothing) = False Then Return "#Error"
            GVARS.G_CURFIGHTER = GetFighterByName(sHero)
            If GVARS.G_CURFIGHTER < 1 Then
                GVARS.G_CURFIGHTER = GetNextFighter(GVARS.G_CURFIGHTER, False)
            End If
        Else
            GVARS.G_CURFIGHTER = GetNextFighter(GVARS.G_CURFIGHTER, True)
        End If

        'событие BattleOnNextEvent
        Dim eventId As Integer = mScript.mainClass(classB).Properties("BattleOnNextEvent").eventId, res As String = ""
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, {prevFighter.ToString, GVARS.G_CURFIGHTER.ToString}, "BattleOnNextEvent", False)
            Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
            If bFinished <> BattleFinishEnum.CUSTOM Then
                If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                Return StopBattle(bFinished)
            End If
            If res = "False" Then
                Return ""
            ElseIf res = "#Error" Then
                Return res
            End If
        End If

        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If GVARS.G_CURFIGHTER < 0 Then
            'все бойцы походили - новый раунд
            MoveFighterBack(hDoc, prevFighter)
            If prevVictim > -1 Then MoveFighterBack(hDoc, prevVictim)
            Return NewTurn()
        End If

        'Подготовка списка возможных жертв
        If PrepareVictimsList(GVARS.G_CURFIGHTER) = "#Error" Then Return "#Error"
        If lstVictims.Count = 0 Then
            Return StopBattle()
        End If

        'убираем предыдущих бойцов
        If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
        If prevFighter > -1 Then
            'Dim life As Double = 0, runAway As Boolean = False
            'If ReadFighterPropertyDbl("Life", prevFighter, life, {prevFighter.ToString}) = False Then Return "#Error"
            'If ReadFighterPropertyBool("RunAway", prevFighter, runAway, {prevFighter.ToString}) = False Then Return "#Error"
            'If life <= 0 OrElse runAway Then
            '    '!!!!!!!!!удаление ячейки с бойцами того же ранга, смещение бойцов, удаление армии!!!!!!!!
            '    'предыдующий боец умер/сбежал - удаляем его html-контейнер
            '    Dim hFighterConvas As HtmlElement = FindFighterHTMLContainer(prevFighter, hDoc)
            '    If IsNothing(hFighterConvas) Then
            '        Dim msFighterConvas As mshtml.IHTMLDOMNode = hFighterConvas.DomElement
            '        msFighterConvas.removeNode(True)
            '    End If
            'Else
            '    'возвращаем бойца на место
            '    MoveFighterBack(hDoc, prevFighter)
            'End If
            'к этому моменту мертвые и убежавшие должны быть убраны
            MoveFighterBack(hDoc, prevFighter)
        End If

        If prevVictim > -1 Then
            MoveFighterBack(hDoc, prevVictim)
        End If
        If blnFirstTime = False Then Wait()

        'перемещаем атакующего
        If MoveFighterToFront(hDoc, GVARS.G_CURFIGHTER, FighterFromLeft, True) = "#Error" Then Return "#Error"
        Wait()

        'выбор типа удара - событие BattleAttackTypeSelect
        Dim isUnderControl As Boolean = IsFighterUnderControl(GVARS.G_CURFIGHTER)
        eventId = mScript.mainClass(classB).Properties("BattleAttackTypeSelect").eventId
        res = mScript.eventRouter.RunEvent(eventId, {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, isUnderControl.ToString}, "BattleAttackTypeSelect", False)
        If res = "#Error" Then Return res

        If Not isUnderControl Then
            'боец неподконтрольный
            If res = "False" Then Return ""
            If GVARS.G_STRIKETYPE < 0 Then GVARS.G_STRIKETYPE = ChooseStrikeType(GVARS.G_CURFIGHTER)
            If GVARS.G_STRIKETYPE = -1 Then Return "#Error"

            Dim delay As Integer = 100
            If ReadPropertyInt(classB, "Delay", -1, -1, delay, Nothing) = False Then Return "#Error"
            If delay < 0 Then delay = 0
            Wait(delay)

            If Hit(GVARS.G_CURFIGHTER, GVARS.G_CURVICTIM, GVARS.G_STRIKETYPE) = "#Error" Then Return "#Error"
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Выжидает указанное количество миллисекунд.
    ''' </summary>
    ''' <param name="waitTime">количество миллисекунд. Если -1, то оперделяется из Battle.FightersSpeed</param>
    Public Sub Wait(Optional ByVal waitTime As Integer = -1)
        If waitTime < 0 Then
            Dim classB As Integer = mScript.mainClassHash("B")
            waitTime = 1000
            ReadPropertyInt(classB, "FightersSpeed", -1, -1, waitTime, Nothing)
        End If

        Dim bWatch As New Stopwatch
        bWatch.Start()
        Do While bWatch.ElapsedMilliseconds <= CLng(waitTime)
            Application.DoEvents()
        Loop
        bWatch.Stop()
    End Sub

    ''' <summary>Определяет управляет ли Игрок бойцом или нет</summary>
    ''' <param name="fId">Id бойца</param>
    Public Function IsFighterUnderControl(ByVal fId As Integer) As Boolean
        Dim fr As Integer = 0, arrs() As String = {fId.ToString}, fEnemy As IsEnemyEnum = IsEnemyEnum.CheckError
        If ReadFighterPropertyInt("Friendliness", fId, fr, arrs) = False Then Return False
        If fr = -1 Then Return True 'главный герой
        fEnemy = IsEnemy(-1, fId)
        If fEnemy = IsEnemyEnum.Enemy Then Return False 'враг

        If Fighters(fId).armyId > -1 Then
            'боец внутри армии
            Dim classArmy As Integer = mScript.mainClassHash("Army")
            Dim UnderControl As Integer = 0
            ReadPropertyInt(classArmy, "UnderControl", Fighters(fId).armyId, -1, UnderControl, {Fighters(fId).armyId})
            If UnderControl = 1 Then
                Return True 'бойцы армии подконтрольные
            ElseIf UnderControl = 2 Then
                Return False 'бойцы армии неподконтрольные
            End If
        End If
        If fEnemy = IsEnemyEnum.Neutral Then Return False 'боец нейтральный - неподконтрольный

        Dim classB As Integer = mScript.mainClassHash("B")
        If fr = 2 Then
            'призванный
            Dim ControlSummoned As Boolean = False
            ReadPropertyBool(classB, "ControlSummoned", -1, -1, ControlSummoned, Nothing)
            Return ControlSummoned
        End If

        'персонаж - друг
        Dim ControlFriends As Boolean = False
        ReadPropertyBool(classB, "ControlFriends", -1, -1, ControlFriends, Nothing)
        Return ControlFriends
    End Function

    ''' <summary>устанавливаем картинку атаки (подразумевается, что GVARS.G_CURFIGHTER уже выбрано)</summary>
    ''' <param name="fId">Id атакующего бойца</param>
    Public Function SetAttackPicture(ByVal fId As Integer) As String
        'устанавливаем картинку атаки
        Dim magicBookId As Integer = Fighters(fId).magicBookId
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim mgArrs() As String = {magicBookId.ToString, GVARS.G_CURMAGIC.ToString}
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        Dim imgVS As HtmlElement = hDoc.GetElementById("vsPicture")
        If IsNothing(imgVS) Then Return "" ' _ERROR("Ошибка в структуре html-документа.")
        If GVARS.G_STRIKETYPE = 0 Then
            'обычная атака
            Dim attackPicture As String = ""
            If ReadFighterProperty("AttackPicture", GVARS.G_CURFIGHTER, attackPicture, {fId.ToString}) = False Then Return "#Error"
            attackPicture = UnWrapString(attackPicture).Replace("\"c, "/"c)
            imgVS.SetAttribute("src", attackPicture)
        ElseIf GVARS.G_STRIKETYPE = 1 Then
            'магическая атака
            If magicBookId < 0 Then Return _ERROR("Магическая атака невозможна. У бойца нет книги магий!")
            If GVARS.G_CURMAGIC < 0 OrElse IsNothing(mScript.mainClass(classMg).ChildProperties(magicBookId)("Name").ThirdLevelProperties) OrElse _
                GVARS.G_CURMAGIC > mScript.mainClass(classMg).ChildProperties(magicBookId)("Name").ThirdLevelProperties.Count - 1 Then Return _
                _ERROR("Заклинания с Id " & GVARS.G_CURMAGIC.ToString & "  в книге магий героя с Id " & GVARS.G_CURFIGHTER.ToString & " не существует!")
            Dim attackPicture As String = ""
            If ReadProperty(classMg, "Picture", magicBookId, GVARS.G_CURMAGIC, attackPicture, mgArrs) = False Then Return "#Error"
            attackPicture = UnWrapString(attackPicture).Replace("\"c, "/"c)
            imgVS.SetAttribute("src", attackPicture)
        Else
            'другая атака
            Dim propName As String = "AttackPictureType" & GVARS.G_STRIKETYPE.ToString
            If Fighters(fId).heroProps.ContainsKey(propName) Then
                Dim attackPicture As String = ""
                If ReadFighterProperty(propName, GVARS.G_CURFIGHTER, attackPicture, {fId.ToString}) = False Then Return "#Error"
                attackPicture = UnWrapString(attackPicture).Replace("\"c, "/"c)
                imgVS.SetAttribute("src", attackPicture)
            End If
        End If
        Return ""
    End Function

    Public Function Hit(ByVal fId As Integer, ByVal vId As Integer, ByVal strikeType As Integer) As String
        '        Dim hStr As String, fVar As Object, aInt As Long, bStruct As Long, hEl As IHTMLElement
        '        Dim calcDamage As Single, pLife As Long, isEnemy As Boolean
        '        Dim ShoudBeVictim As Boolean, aBool As Boolean, i As Long
        '        bStruct = FindStruct("B")
        '        isEnemy = Battle_IsEnemy(G_VAR.G_CURFIGHTER)

        '        hStr = ""
        '        hEl = csWb.hdocLoc.getElementById("battle_vs")
        If vId > -1 Then
            If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
        End If
        GVARS.G_STRIKETYPE = strikeType
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document, magicBookId As Integer = Fighters(fId).magicBookId, eventId As Integer = 0
        Dim classMg As Integer = mScript.mainClassHash("Mg"), classB As Integer = mScript.mainClassHash("B"), mgArrs() As String = {magicBookId.ToString, GVARS.G_CURMAGIC.ToString}
        Dim calcDamage As Double = 0 'расчетное повреждение
        Dim res As String = "", attackRefused As Boolean = False

        'устанавливаем картинку атаки
        If SetAttackPicture(fId) = "#Error" Then Return "#Error"

        Dim ShoudBeVictim As Boolean = True
        If GVARS.G_STRIKETYPE = 1 Then
            Dim mType As MagicTypesEnum = MagicTypesEnum.Damage
            If ReadPropertyInt(classMg, "MagicType", magicBookId, GVARS.G_CURMAGIC, mType, mgArrs) = False Then Return "#Error"
            If mType = MagicTypesEnum.FriendsEnhancer OrElse mType = MagicTypesEnum.OtherFriend Then
                If PrepareVictimsList(GVARS.G_CURFIGHTER, True) = False Then Return "#Error"
                If lstVictims.Count = 0 Then Return _ERROR("Использование магии " & mScript.mainClass(classMg).ChildProperties(magicBookId)("Name").ThirdLevelProperties(GVARS.G_CURMAGIC) & " не возможно! В битве нет ни одного вашего союзника!")
                ShoudBeVictim = True
            ElseIf mType = MagicTypesEnum.Damage OrElse mType = MagicTypesEnum.OtherWithvictim Then
                ShoudBeVictim = True
            Else
                ShoudBeVictim = False
            End If
        End If

        If ShoudBeVictim Then
            'тип атаки требует выбора жертвы
            If vId < 0 Then
                'Эта часть блока данной функции будет выполнена только если она запущена без указания Id жертвы vId.
                'Содержит данный блок код выбора жертвы. В случае, если выбор предоставляется Игроку с помощью действий, то произойдет выход из функции.
                'При выборе действия с Battle.HitVictim данная функция будет запущена еще раз, но, если жертва была выбрана, то данный блок будет пропущен
                'Событие BattleOnVictimSelect
                Dim isUnderControl As Boolean = IsFighterUnderControl(fId)
                If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
                eventId = mScript.mainClass(classB).Properties("BattleOnVictimSelect").eventId
                res = ""
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, {GVARS.G_CURFIGHTER.ToString, lstVictims.Count, isUnderControl.ToString}, "BattleOnVictimSelect", False)
                    If res = "#Error" Then
                        Return res
                    ElseIf res = "False" Then
                        'Остановка скрипта BattleOnVictimSelect
                        Return ""
                    End If
                End If

                'выбор жертвы
                If isUnderControl Then
                    'для подконтрольных бойцов
                    If lstVictims.Count = 1 Then
                        GVARS.G_CURVICTIM = lstVictims(0)
                    Else
                        Return "" 'остановка так как не выбрана жертва (это должно произойти при выборе действия, описанного в BattleOnVictimSelect)
                    End If
                Else
                    'для НЕподконтрольных бойцов - автовыбор жертвы
                    GVARS.G_CURVICTIM = ChooseVictim(GVARS.G_CURFIGHTER)
                End If
            Else
                GVARS.G_CURVICTIM = vId
            End If

            If GVARS.G_CURVICTIM < 0 Then Return _ERROR("В событии BattleOnVictimSelect не был произведен выбор жертвы.")
            'перемещаем жертву на фронт
            If MoveFighterToFront(hDoc, GVARS.G_CURVICTIM, Not FighterFromLeft) = "#Error" Then Return "#Error"

            Dim life As Double = 0
            If ReadFighterPropertyDbl("Life", GVARS.G_CURVICTIM, life, {GVARS.G_CURVICTIM.ToString}) = False Then Return "#Error"
            calcDamage = GetDamage(GVARS.G_CURFIGHTER, GVARS.G_CURVICTIM, GVARS.G_STRIKETYPE)
            Wait()
        Else
            'тип атаки не требует выбора жертвы
            GVARS.G_CURVICTIM = -1
        End If

        'Событие BattleOnAttackEvent
        eventId = mScript.mainClass(classB).Properties("BattleOnAttackEvent").eventId
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, calcDamage.ToString(provider_points), _
                                                                       magicBookId.ToString, GVARS.G_CURMAGIC.ToString}, "BattleOnAttackEvent", False)
            If res = "#Error" Then Return res
            'проверка на окончание боя
            Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
            If bFinished <> BattleFinishEnum.CUSTOM Then
                If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                Return StopBattle(bFinished)
            End If
            'учет возвращаемого значения
            Dim dam As Double = 0
            If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam) Then
                calcDamage = dam
            ElseIf res = "False" Then
                attackRefused = True
                GoTo skip_events
            End If
        End If

        Dim classH As Integer = mScript.mainClassHash("H")
        If GVARS.G_CURVICTIM >= 0 Then
            'Событие HeroDefenceEvent
            For i As Integer = 1 To 2
                Dim hArrs() As String = {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, calcDamage.ToString(provider_points), GVARS.G_CURMAGIC.ToString}
                If i = 1 Then
                    'глобальное
                    eventId = mScript.mainClass(classH).Properties("HeroDefenceEvent").eventId
                Else
                    'данной жертвы
                    eventId = Fighters(GVARS.G_CURVICTIM).heroProps("HeroDefenceEvent").eventId
                End If
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, hArrs, "HeroDefenceEvent", False)
                    If res = "#Error" Then Return res
                    'проверка на окончание боя
                    Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
                    If bFinished <> BattleFinishEnum.CUSTOM Then
                        If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                        Return StopBattle(bFinished)
                    End If
                    'учет возвращаемого значения
                    Dim dam As Double = 0
                    If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam) Then
                        calcDamage = dam
                    ElseIf res = "False" Then
                        attackRefused = True
                        GoTo skip_events
                    End If
                End If
            Next i
        End If

        'Событие HeroAttackEvent
        For i As Integer = 1 To 2
            Dim hArrs() As String = {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, calcDamage.ToString(provider_points), GVARS.G_CURMAGIC.ToString}
            If i = 1 Then
                'глобальное
                eventId = mScript.mainClass(classH).Properties("HeroAttackEvent").eventId
            Else
                'данного бойца
                eventId = Fighters(GVARS.G_CURFIGHTER).heroProps("HeroAttackEvent").eventId
            End If
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, hArrs, "HeroAttackEvent", False)
                If res = "#Error" Then Return res
                'проверка на окончание боя
                Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
                If bFinished <> BattleFinishEnum.CUSTOM Then
                    If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                    Return StopBattle(bFinished)
                End If
                'учет возвращаемого значения
                Dim dam As Double = 0
                If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam) Then
                    calcDamage = dam
                ElseIf res = "False" Then
                    attackRefused = True
                    GoTo skip_events
                End If
            End If
        Next i

        Dim classAb As Integer = mScript.mainClassHash("Ab")
        'События AbilityDefenceEvent
        Dim abSetId As Integer = -1
        If GVARS.G_CURVICTIM >= 0 AndAlso abSetId >= 0 AndAlso IsNothing(mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties) = False _
            AndAlso mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties.Count > 0 Then
            abSetId = Fighters(GVARS.G_CURVICTIM).abilitySetId
            Dim globalEventId = mScript.mainClass(classAb).Properties("AbilityDefenceEvent").eventId
            Dim abSetEventId = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityDefenceEvent").eventId
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityDefenceEvent")
            For abId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                Dim wasEvent As Boolean = False
                For i As Integer = 1 To 3
                    Dim hArrs() As String = {abSetId.ToString, abId.ToString, GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, GVARS.G_CURMAGIC.ToString, calcDamage.ToString(provider_points)}
                    If i = 1 Then
                        'глобальное
                        eventId = globalEventId
                    ElseIf i = 2 Then
                        'данного набора
                        eventId = abSetEventId
                    Else
                        'данного бойца
                        eventId = ch.ThirdLevelEventId(abId)
                    End If
                    If eventId > 0 Then
                        wasEvent = True
                        res = mScript.eventRouter.RunEvent(eventId, hArrs, "AbilityDefenceEvent", False)
                        If res = "#Error" Then Return res
                        'проверка на окончание боя
                        Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
                        If bFinished <> BattleFinishEnum.CUSTOM Then
                            If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                            Return StopBattle(bFinished)
                        End If
                        'учет возвращаемого значения
                        Dim dam As Double = 0
                        If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam) Then
                            calcDamage = dam
                        ElseIf res = "False" Then
                            attackRefused = True
                            GoTo skip_events
                        End If
                    End If
                Next i

                If wasEvent Then
                    'проигрывание звука
                    Dim abSound As String = ""
                    If ReadProperty(classAb, "SoundOnApply", abSetId, abId, abSound, {abSetId.ToString, abId.ToString}) = False Then Return "#Error"
                    abSound = UnWrapString(abSound)
                    If String.IsNullOrEmpty(abSound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, abSound)) Then
                        Say(abSound, 100, False)
                    End If
                End If
            Next abId
        End If

        'События AbilityAttackEvent
        abSetId = Fighters(GVARS.G_CURFIGHTER).abilitySetId
        If abSetId >= 0 AndAlso IsNothing(mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties) = False _
            AndAlso mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties.Count > 0 Then
            Dim globalEventId = mScript.mainClass(classAb).Properties("AbilityAttackEvent").eventId
            Dim abSetEventId = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityAttackEvent").eventId
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityAttackEvent")
            For abId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                Dim wasEvent As Boolean = False
                For i As Integer = 1 To 3
                    Dim hArrs() As String = {abSetId.ToString, abId.ToString, GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, GVARS.G_CURMAGIC.ToString, calcDamage.ToString(provider_points)}
                    If i = 1 Then
                        'глобальное
                        eventId = globalEventId
                    ElseIf i = 2 Then
                        'данного набора
                        eventId = abSetEventId
                    Else
                        'данного бойца
                        eventId = ch.ThirdLevelEventId(abId)
                    End If
                    If eventId > 0 Then
                        wasEvent = True
                        res = mScript.eventRouter.RunEvent(eventId, hArrs, "AbilityAttackEvent", False)
                        If res = "#Error" Then Return res
                        'проверка на окончание боя
                        Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
                        If bFinished <> BattleFinishEnum.CUSTOM Then
                            If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                            Return StopBattle(bFinished)
                        End If
                        'учет возвращаемого значения
                        Dim dam As Double = 0
                        If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam) Then
                            calcDamage = dam
                        ElseIf res = "False" Then
                            attackRefused = True
                            GoTo skip_events
                        End If
                    End If
                Next i

                If wasEvent Then
                    'проигрывание звука
                    Dim abSound As String = ""
                    If ReadProperty(classAb, "SoundOnApply", abSetId, abId, abSound, {abSetId.ToString, abId.ToString}) = False Then Return "#Error"
                    abSound = UnWrapString(abSound)
                    If String.IsNullOrEmpty(abSound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, abSound)) Then
                        Say(abSound, 100, False)
                    End If
                End If
            Next abId
        End If

        'пауза
        Dim delay As Integer = 100
        If ReadPropertyInt(classB, "Delay", -1, -1, delay, Nothing) = False Then Return "#Error"
        If delay < 0 Then delay = 100
        Wait(delay)

        'эффект удара
        '        If Len(hEff) > 0 Then
        '            hPar1 = modWork.PrepareStrToPrint(hPar1)
        '            hPar2 = modWork.PrepareStrToPrint(hPar2)
        '            bStr = "applyFilterEx(document.all[" + CStr(hEl.sourceIndex) + "], '" + hEff + "', '" + hDur + "', '" + hPar1 + "', '" + hPar2 + "');"
        '            csWb.hdocLoc.parentWindow.execScript bStr
        '            hEl.className = hClass
        '            bStr = "playFilterEx(document.all[" + CStr(hEl.sourceIndex) + "]);"
        '            csWb.hdocLoc.parentWindow.execScript bStr
        '        Else
        '            hEl.className = hClass
        '        End If

        If GVARS.G_STRIKETYPE = 1 Then
            'удар магический
            mgArrs = {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, calcDamage.ToString(provider_points), magicBookId.ToString, GVARS.G_CURMAGIC.ToString}

            'сначала - звук
            Dim snd As String = ""
            If ReadProperty(classMg, "Sound", magicBookId, GVARS.G_CURMAGIC, snd, {magicBookId.ToString, GVARS.G_CURMAGIC.ToString}) = False Then Return False
            snd = UnWrapString(snd)
            If String.IsNullOrEmpty(snd) = False Then
                If Say(snd, 100, False) = "#Error" Then Return "#Error"
            End If

            'Событие MagicCastEvent
            For i As Integer = 1 To 3
                If i = 1 Then
                    'глобальное
                    eventId = mScript.mainClass(classMg).Properties("MagicCastEvent").eventId
                ElseIf i = 2 Then
                    'данной книги магий
                    eventId = mScript.mainClass(classMg).ChildProperties(magicBookId)("MagicCastEvent").eventId
                Else
                    'данной магии
                    eventId = mScript.mainClass(classMg).ChildProperties(magicBookId)("MagicCastEvent").ThirdLevelEventId(GVARS.G_CURMAGIC)
                End If
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, mgArrs, "MagicCastEvent", False)
                    If res = "#Error" Then Return False
                End If

                'проверка на окончание боя
                Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
                If bFinished <> BattleFinishEnum.CUSTOM Then
                    If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                    Return StopBattle(bFinished)
                End If
            Next i

        ElseIf GVARS.G_STRIKETYPE = 0 Then
            'обычный удар
            'проигрываем звук при атаке Hero.SoundOnAttack
            Dim sound As String = ""
            If ReadFighterProperty("SoundOnAttack", GVARS.G_CURFIGHTER, sound, {GVARS.G_CURFIGHTER.ToString}) = False Then Return "#Error"
            sound = UnWrapString(sound)
            If String.IsNullOrEmpty(sound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, sound)) Then
                Say(sound, 100, False)
                Wait(100)
            End If
        Else
            'спецудар
            'проигрываем звук при атаке Hero.SoundOnAttackTypeX
            Dim sound As String = "", propName As String = "SoundOnAttackType" & GVARS.G_STRIKETYPE.ToString
            If Fighters(GVARS.G_CURFIGHTER).heroProps.ContainsKey(propName) Then
                If ReadFighterProperty(propName, GVARS.G_CURFIGHTER, sound, {GVARS.G_CURFIGHTER.ToString}) = False Then Return "#Error"
                sound = UnWrapString(sound)
                If String.IsNullOrEmpty(sound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, sound)) Then
                    Say(sound, 100, False)
                    Wait(100)
                End If
            End If
        End If

        If calcDamage <> 0 Then
            'наносим указанный урон жизни жервы
            Dim life As Double = 0
            If ReadFighterPropertyDbl("Life", GVARS.G_CURVICTIM, life, {GVARS.G_CURVICTIM.ToString}) = False Then Return "#Error"
            life -= calcDamage
            If PropertiesRouter(classH, "Life", {GVARS.G_CURVICTIM.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, life.ToString(provider_points)) = "#Error" Then Return "#Error"

            'проверка на окончание боя
            Dim bFinished As BattleFinishEnum = CheckForBattleFinished()
            If bFinished <> BattleFinishEnum.CUSTOM Then
                'окончание боя
                If bFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
                Return StopBattle(bFinished)
            ElseIf calcDamage > 0 Then
                'звук Hero.SoundOnInjured 
                Dim sound As String = ""
                If ReadFighterProperty("SoundOnInjured", GVARS.G_CURVICTIM, sound, {GVARS.G_CURVICTIM.ToString}) = False Then Return "#Error"
                sound = UnWrapString(sound)
                If String.IsNullOrEmpty(sound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, sound)) Then
                    Say(sound, 100, False)
                    Wait(100)
                End If
            End If
        End If

        GVARS.G_CANUSEMAGIC = False
        Application.DoEvents()
skip_events:

        'проверка на окончание боя
        Dim btlFinished As BattleFinishEnum = CheckForBattleFinished()
        If btlFinished <> BattleFinishEnum.CUSTOM Then
            If btlFinished = BattleFinishEnum.CHECK_ERROR Then Return "#Error"
            Return StopBattle(btlFinished)
        End If

        If attackRefused Then
            'звук Hero.SoundOnFightOff
            Dim sound As String = ""
            If ReadFighterProperty("SoundOnFightOff", GVARS.G_CURVICTIM, sound, {GVARS.G_CURVICTIM.ToString}) = False Then Return "#Error"
            sound = UnWrapString(sound)
            If String.IsNullOrEmpty(sound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, sound)) Then
                Say(sound, 100, False)
                Wait(100)
            End If
        End If

        'Событие BattleAfterAttackEvent
        eventId = mScript.mainClass(classB).Properties("BattleAfterAttackEvent").eventId
        If eventId > 0 Then
            If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
            Dim arrs() = {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, GVARS.G_STRIKETYPE.ToString, calcDamage.ToString(provider_points), magicBookId.ToString, GVARS.G_CURMAGIC.ToString, (attackRefused = False).ToString}
            res = mScript.eventRouter.RunEvent(eventId, arrs, "BattleAfterAttackEvent", False)
            If res = "#Error" Then Return "#Error"
        Else
            'событие пустое - просто следующий ход
            If NextTurn() = "#Error" Then Return "#Error"
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Расчитывает повреждение при любой атаке, кроме магической
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="vId">Id жертвы</param>
    ''' <param name="strikeType"></param>
    ''' <returns>расчетное повреждение исходя из Battle.DamageCalculation</returns>
    Private Function GetDamage(ByVal fId As Integer, ByVal vId As Integer, ByVal strikeType As Integer) As Double
        If strikeType = 1 Then Return 0
        Dim classB As Integer = mScript.mainClassHash("B")
        Dim eventId As Integer = mScript.mainClass(classB).Properties("DamageCalculation").eventId, res As String = ""
        If eventId < 1 Then Return mScript.ExecuteString({"?C.Round(H[B.CurFighter].Power - H[B.CurVictim].Defence / 2, 0)"}, Nothing)

        res = mScript.eventRouter.RunEvent(eventId, {fId.ToString, vId.ToString, strikeType.ToString}, "DamageCalculation", False)
        Dim dam As Double = 0
        Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, dam)
        Return dam
    End Function

    ''' <summary>
    ''' Рандомизирует повреждение от атаки, чтобы оно не было четко абсолютно одинаковым при каждом ударе. На выходе получается целое число
    ''' </summary>
    ''' <param name="calcDamage">расчетое повреждение</param>
    ''' <param name="maxValue">максимум</param>
    Public Function RandomizeDamage(ByVal calcDamage As Double, Optional maxValue As Double = Double.MaxValue, Optional digits As Integer = 0) As Double
        Dim classB As Integer = mScript.mainClassHash("B")
        'получаем разброс и минимальное значение
        Dim RandomDispertion As String = "", MinimalDamage As Double = 1
        If ReadProperty(classB, "RandomDispertion", -1, -1, RandomDispertion, Nothing) = False Then Return 0
        RandomDispertion = UnWrapString(RandomDispertion)
        If String.IsNullOrEmpty(RandomDispertion) Then RandomDispertion = "100%"
        If ReadPropertyDbl(classB, "MinimalDamage", -1, -1, MinimalDamage, Nothing) = False Then Return 0

        'расчитываем амплитуду, в пределах которой возможны значения повреждения
        Dim amplitude As Double
        If RandomDispertion.EndsWith("%") Then
            Dim percents As Double = Val(RandomDispertion.Substring(0, RandomDispertion.Length - 1))
            amplitude = calcDamage * percents / 100
        Else
            amplitude = Val(RandomDispertion)
        End If

        'расчитываем итоговое повреждение
        calcDamage = Math.Min(calcDamage + Rnd() * amplitude * 2 - amplitude, maxValue)
        calcDamage = Math.Max(calcDamage, MinimalDamage)

        Return Math.Round(calcDamage, digits)
    End Function

    ''' <summary>
    ''' Создаем список бойцов, доступных для атаки/воздействия
    ''' </summary>
    ''' <param name="fighterId">Id текущего бойца</param>
    ''' <param name="blnGetFriends">Воздействие на дружественных персонажей или на врагов</param>
    Public Function PrepareVictimsList(ByVal fighterId As Long, Optional blnGetFriends As Boolean = False) As String
        lstVictims.Clear()
        Dim pLife As Double = 0, runAway As Boolean = False, army As String = "", armyId As Integer = -1
        Dim classArmy As Integer = mScript.mainClassHash("Army")
        If ReadFighterProperty("Army", fighterId, army) = False Then Return "#Error"
        armyId = GetSecondChildIdByName(army, mScript.mainClass(classArmy).ChildProperties)

        Dim globalEventId As Integer = -1, armyEventId As Integer = -1
        If armyId > -1 Then
            globalEventId = mScript.mainClass(classArmy).Properties("ArmyOnHostilityChecking").eventId
            armyEventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnHostilityChecking").eventId
        End If

        For vId As Integer = 0 To Fighters.Count - 1
            If vId = fighterId Then Continue For
            Dim arrs() As String = {vId.ToString}
            'герой мертв - пропускаем
            If ReadFighterPropertyDbl("Life", vId, pLife, arrs) = False Then Return "#Error"
            If pLife <= 0 Then Continue For
            'герой сбежал - пропускаем
            If ReadFighterPropertyBool("RunAway", vId, runAway, arrs) = False Then Return "#Error"
            If runAway Then Continue For
            'проверка на враждебность
            Dim fEnemy As IsEnemyEnum = IsEnemy(fighterId, vId)
            If fEnemy = IsEnemyEnum.CheckError Then
                Return "#Error"
            ElseIf fEnemy = IsEnemyEnum.Neutral OrElse (blnGetFriends AndAlso fEnemy = IsEnemyEnum.Enemy) OrElse (blnGetFriends = False AndAlso fEnemy = IsEnemyEnum.Friendly) Then
                Continue For
            End If
            'все проверки пройдены - герой включается в список
            lstVictims.Add(vId)
        Next vId
        Return ""
    End Function

    ''' <summary>
    ''' Определяет враждебность бойца checkingFighterId по отношению к curFighterId
    ''' </summary>
    ''' <param name="curFighterId">Id текущего бойца. Если -1,то по отношению к главному герою</param>
    ''' <param name="checkingFighterId">Id проверяемого бойца</param>
    Public Function IsEnemy(ByVal curFighterId As Integer, ByVal checkingFighterId As Integer) As IsEnemyEnum
        Dim classArmy As Integer = mScript.mainClassHash("Army")

        'Получаем свойства Friendliness
        Dim curFigherFr As Integer = -1, checkFighterFr As Integer = -1
        If curFighterId = -1 Then
            curFigherFr = -1
        Else
            If ReadFighterPropertyInt("Friendliness", curFighterId, curFigherFr, {curFighterId.ToString}) = False Then Return IsEnemyEnum.CheckError
        End If
        If ReadFighterPropertyInt("Friendliness", checkingFighterId, checkFighterFr, {checkFighterFr.ToString}) = False Then Return IsEnemyEnum.CheckError
        'если бойцы призванные, то получаем вместо них хозяев
        If curFigherFr = 2 Then
            curFighterId = GetSummonOwner(curFighterId)
            If ReadFighterPropertyInt("Friendliness", curFighterId, curFigherFr, {curFighterId.ToString}) = False Then Return IsEnemyEnum.CheckError
        End If
        If checkFighterFr = 2 Then
            checkingFighterId = GetSummonOwner(checkingFighterId)
            If ReadFighterPropertyInt("Friendliness", checkingFighterId, checkFighterFr, {checkFighterFr.ToString}) = False Then Return IsEnemyEnum.CheckError
        End If
        'сейчас бойцы гарантированно не призванные

        If curFighterId = -1 Then
            'проверка враждебности по отношению к главному герою. Поиск главного героя
            For i As Integer = 0 To Fighters.Count - 1
                Dim fr As Integer = 0
                If ReadFighterPropertyInt("Friendliness", i, fr, {i.ToString}) = False Then Return IsEnemyEnum.CheckError
                If fr = -1 Then
                    curFighterId = i
                    Exit For
                End If
            Next i

            If curFighterId = -1 Then
                'главный герой не учавствует в бою. Ищем первого попавшегося друга
                For i As Integer = 0 To Fighters.Count - 1
                    Dim fr As Integer = 0
                    If ReadFighterPropertyInt("Friendliness", i, fr, {i.ToString}) = False Then Return IsEnemyEnum.CheckError
                    If fr = 1 Then
                        curFighterId = i
                        Exit For
                    End If
                Next i
            End If
        End If

        'событие ArmyOnHostilityChecking проверки враждебности бойца
        If curFighterId > -1 AndAlso Fighters(curFighterId).armyId > -1 Then
            Dim res As String = ""
            Dim arrs() As String = {Fighters(curFighterId).armyId.ToString, curFighterId.ToString, Fighters(checkingFighterId).armyId.ToString, checkingFighterId.ToString}
            'событие ArmyOnHostilityChecking
            Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnHostilityChecking").eventId
            If eventId > 0 Then
                'глобальное
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnHostilityChecking", False)
                If res = "True" Then
                    Return IsEnemyEnum.Enemy
                ElseIf res = "False" Then
                    Return IsEnemyEnum.Friendly
                ElseIf res = "0" Then
                    Return IsEnemyEnum.Neutral
                ElseIf res = "#Error" Then
                    Return IsEnemyEnum.CheckError
                End If
            End If

            eventId = mScript.mainClass(classArmy).ChildProperties(Fighters(curFighterId).armyId)("ArmyOnHostilityChecking").eventId
            If eventId > 0 Then
                'данной армии
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnHostilityChecking", False)
                If res = "True" Then
                    Return IsEnemyEnum.Enemy
                ElseIf res = "False" Then
                    Return IsEnemyEnum.Friendly
                ElseIf res = "0" Then
                    Return IsEnemyEnum.Neutral
                ElseIf res = "#Error" Then
                    Return IsEnemyEnum.CheckError
                End If
            End If

            eventId = -1
            If Fighters(curFighterId).armyUnitId > -1 Then eventId = mScript.mainClass(classArmy).ChildProperties(Fighters(curFighterId).armyId)("ArmyOnHostilityChecking").ThirdLevelEventId(Fighters(curFighterId).armyUnitId)
            If eventId > -1 Then
                'данного юнита
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnHostilityChecking", False)
                If res = "True" Then
                    Return IsEnemyEnum.Enemy
                ElseIf res = "False" Then
                    Return IsEnemyEnum.Friendly
                ElseIf res = "0" Then
                    Return IsEnemyEnum.Neutral
                ElseIf res = "#Error" Then
                    Return IsEnemyEnum.CheckError
                End If
            End If
        End If

        'события не описано или текущий герой вне армии - проверка свойства Friendliness
        If curFigherFr = -1 OrElse curFigherFr = 1 Then
            If checkFighterFr = 0 Then
                Return IsEnemyEnum.Enemy
            Else
                Return IsEnemyEnum.Friendly
            End If
        Else
            If checkFighterFr = 0 Then
                Return IsEnemyEnum.Friendly
            Else
                Return IsEnemyEnum.Enemy
            End If
        End If
    End Function

    ''' <summary>
    ''' Получает Id хозяина призванного. Если хозяина нет, то -1
    ''' </summary>
    ''' <param name="fighterId">Id призванного</param>
    ''' <param name="blnMainOwner">получить главного хозяина</param>
    Public Function GetSummonOwner(ByVal fighterId As Integer, Optional blnMainOwner As Boolean = True) As Integer
        Dim owner As String = "", ownerId As Integer = -1

        If ReadFighterProperty("HeroOwner", fighterId, owner, {fighterId.ToString}) = False Then Return -1
        ownerId = GetFighterByName(owner)
        If Not blnMainOwner OrElse ownerId = -1 Then Return ownerId

        'получаем главного хозяина
        Do
            fighterId = ownerId
            If ReadFighterProperty("HeroOwner", fighterId, owner, {fighterId.ToString}) = False Then Return -1
            ownerId = GetFighterByName(owner)
            If ownerId = -1 Then Return fighterId
        Loop
    End Function

    ''' <summary>
    ''' Убивает призванных бойцов, хозяином которых является указанный персонаж
    ''' </summary>
    ''' <param name="hostId">Id хозяина</param>
    ''' <returns>Количество убитых</returns>
    Public Function KillSummoned(ByVal hostId As Integer) As Integer
        Dim shoudKill As Boolean = False, classB As Integer = mScript.mainClassHash("B"), classH As Integer = mScript.mainClassHash("H")
        If ReadPropertyBool(classB, "KillSummonedWidthOwner", -1, -1, shoudKill, Nothing) = False Then Return "#Error"
        If Not shoudKill Then Return ""

        Static killed As Integer = 0
        If killingSummoned = False Then killed = 0
        Dim kPrev As Boolean = killingSummoned
        killingSummoned = True

        For fId As Integer = 0 To Fighters.Count - 1
            Dim fr As Integer = 0
            If ReadFighterPropertyInt("Friendliness", fId, fr, {fId.ToString}) = False Then Return "#Error"
            If fr <> 2 Then Continue For 'не призваный
            Dim ownerId As Integer = GetSummonOwner(fId)
            If ownerId <> hostId Then Continue For

            'призванный принадлежит бойцу hostId
            If PropertiesRouter(classH, "Life", {fId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return "#Error"
            killed += 1
        Next fId

        killingSummoned = kPrev
        Return killed
    End Function

    ''' <summary>
    ''' Функция выбора жертвы компьютером на основании свойства атакующего перснонажа Hero.VictimChoice
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    ''' <returns>Id жертвы</returns>
    Public Function ChooseVictim(ByVal fId As Integer) As Integer
        If lstVictims.Count = 1 Then Return lstVictims(0)
        Dim victimChoice As victimChoiceEnum = victimChoiceEnum.Randomly
        Dim arrs() As String = {fId.ToString}, classB As Integer = mScript.mainClassHash("B")
        ReadFighterPropertyInt("VictimChoice", fId, victimChoice, arrs)

        Dim lstChosen As New List(Of Integer)
        Select Case victimChoice
            Case victimChoiceEnum.MainHeroOnly
                'только ГГ
                For i As Integer = 0 To lstVictims.Count - 1
                    Dim playerType As Integer = 0
                    ReadFighterPropertyInt("Friendliness", lstVictims(i), playerType, {lstVictims(i).ToString})
                    If playerType = -1 Then lstChosen.Add(lstVictims(i))
                Next
            Case victimChoiceEnum.NotMainHero
                'любого, кроме ГГ
                For i As Integer = 0 To lstVictims.Count - 1
                    Dim playerType As Integer = 0
                    ReadFighterPropertyInt("Friendliness", lstVictims(i), playerType, {lstVictims(i).ToString})
                    If playerType <> -1 Then lstChosen.Add(lstVictims(i))
                Next
            Case victimChoiceEnum.TheStrongest
                'выбор самого сильного на основании события Battle.PowerRange
                Dim eventId As Integer = mScript.mainClass(classB).Properties("PowerRange").eventId
                If eventId > 0 Then
                    'если событие не описано - пропускаем (далее будет просто случайный выбор)
                    Dim lstRanges As New List(Of Double), maxRange As Double = Double.MinValue, range As Double = 0, res As String = ""
                    'составляем список бойцов с наиболее высоким рангом
                    For i As Integer = 0 To lstVictims.Count - 1
                        res = mScript.eventRouter.RunEvent(eventId, {lstVictims(i)}, "PowerRange", False)
                        If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, range) Then
                            If range >= maxRange Then
                                'новый боец имеет ранг выше или равный предыдущим - сохраняем его
                                maxRange = res
                                lstChosen.Add(lstVictims(i))
                                lstRanges.Add(range)
                            End If
                        ElseIf res = "#Error" Then
                            Exit For
                        End If
                    Next i
                    'на данный момент в списке lstChosen могут быть лишние персонажи. Удаляем их
                    For i As Integer = lstRanges.Count - 1 To 0 Step -1
                        If lstRanges(i) < maxRange Then lstChosen.RemoveAt(i)
                    Next
                End If
            Case victimChoiceEnum.TheWeakest
                'выбор самого слабого на основании события Battle.PowerRange
                Dim eventId As Integer = mScript.mainClass(classB).Properties("PowerRange").eventId
                If eventId > 0 Then
                    'если событие не описано - пропускаем (далее будет просто случайный выбор)
                    Dim lstRanges As New List(Of Double), minRange As Double = Double.MaxValue, range As Double = 0, res As String = ""
                    'составляем список бойцов с наиболее высоким рангом
                    For i As Integer = 0 To lstVictims.Count - 1
                        res = mScript.eventRouter.RunEvent(eventId, {lstVictims(i)}, "PowerRange", False)
                        If Double.TryParse(res, System.Globalization.NumberStyles.Float, provider_points, range) Then
                            If range <= minRange Then
                                'новый боец имеет ранг выше или равный предыдущим - сохраняем его
                                minRange = res
                                lstChosen.Add(lstVictims(i))
                                lstRanges.Add(range)
                            End If
                        ElseIf res = "#Error" Then
                            Exit For
                        End If
                    Next i
                    'на данный момент в списке lstChosen могут быть лишние персонажи. Удаляем их
                    For i As Integer = lstRanges.Count - 1 To 0 Step -1
                        If lstRanges(i) > minRange Then lstChosen.RemoveAt(i)
                    Next
                End If
        End Select

        If lstChosen.Count = 1 Then
            Return lstChosen(0)
        ElseIf lstChosen.Count > 1 Then
            Return lstChosen(Math.Round(Rnd() * lstChosen.Count - 0.5))
        Else
            'victimChoiceEnum.Randomly
            Return lstVictims(Math.Round(Rnd() * lstVictims.Count - 0.5))
        End If
    End Function

    ''' <summary>
    ''' Автоматически определяет тип удара для бойцов, управляемых компьютером. Это может быть только обычный удар или использование магии
    ''' </summary>
    ''' <param name="fId">Id бойца</param>
    Public Function ChooseStrikeType(ByVal fId As Integer) As Integer
        If Fighters(fId).magicBookId < 0 OrElse HasEnabledMagics(fId) = False Then Return 0 'магии нет - значит обычный удар
        Dim classB As Integer = mScript.mainClassHash("B"), classMg As Integer = mScript.mainClassHash("Mg")

        Dim turnsCount As Integer = 0, arrs() As String = {fId.ToString}
        If ReadPropertyInt(classB, "TurnsCount", -1, -1, turnsCount, arrs) = False Then Return -1

        If turnsCount <= 1 Then
            'FirstRoundMagic, MagicFirstRoundChance
            Dim FirstRoundMagic As String = "", FirstRoundMagicId As Integer = -1, mBookId As Integer = Fighters(fId).magicBookId
            If ReadFighterProperty("FirstRoundMagic", fId, FirstRoundMagic, Nothing) = False Then Return -1
            FirstRoundMagicId = GetThirdChildIdByName(FirstRoundMagic, mBookId, mScript.mainClass(classMg).ChildProperties)
            If FirstRoundMagic > -1 Then
                Dim MagicFirstRoundChance As Integer = 0
                If ReadFighterPropertyInt("MagicFirstRoundChance", fId, MagicFirstRoundChance, Nothing) = False Then Return -1
                If MagicFirstRoundChance <= 0 OrElse Rnd() * 100 < 100 - MagicFirstRoundChance Then
                    Return 0
                Else
                    GVARS.G_CURMAGIC = FirstRoundMagicId
                    Return 1
                End If
            End If
        End If

        Dim MagicChanceTotal As Integer = 0
        If ReadFighterPropertyInt("MagicChanceTotal", fId, MagicChanceTotal, Nothing) = False Then Return -1
        If MagicChanceTotal <= 0 OrElse Rnd() * 100 > MagicChanceTotal Then Return 0

        'Удар магический. Выбираем магию

        'составляем список магий по типам и список весов каждого типа магии
        Dim lstMagics As New SortedList(Of MagicTypesEnum, List(Of Integer)), magicWeights As New SortedList(Of MagicTypesEnum, Double), weight As Double = 0, weightsSum As Double = 0
        'список весов
        If ReadFighterPropertyDbl("MagicChanceDamage", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.Damage, weight)
        If ReadFighterPropertyDbl("MagicChanceSelfEnhancer", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.SelfEnhancer, weight)
        If ReadFighterPropertyDbl("MagicChanceFriendsEnhancer", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.FriendsEnhancer, weight)
        If ReadFighterPropertyDbl("MagicChanceTotalDamage", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.TotalDamage, weight)
        If ReadFighterPropertyDbl("MagicChanceSummoning", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.Summoning, weight)
        If ReadFighterPropertyDbl("MagicChanceOtherWithVictim", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.OtherWithvictim, weight)
        If ReadFighterPropertyDbl("MagicChanceOtherWithoutVictim", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.OtherWithoutVictim, weight)
        If ReadFighterPropertyDbl("MagicChanceOtherWithFriend", fId, weight, arrs) = False Then Return "#Error"
        magicWeights.Add(MagicTypesEnum.OtherFriend, weight)
        weightsSum = magicWeights.Values.Sum
        If weightsSum <= 0 Then Return 0 'шанс всех типов магий 0 - обычный удар

        'список магий
        Dim magicBookId As Integer = Fighters(fId).magicBookId
        Dim magicType As MagicTypesEnum = MagicTypesEnum.Damage
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMg).ChildProperties(magicBookId)("MagicType")
        For mgId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
            If CanUseMagic(fId, magicBookId, mgId) = False Then Continue For
            If ReadPropertyInt(classMg, "MagicType", magicBookId, mgId, magicType, {magicBookId.ToString, mgId.ToString}) = False Then Return -1
            If magicWeights.ContainsKey(magicType) = False OrElse magicWeights(magicType) <= 0 Then Continue For 'если шанс нулевой, то магию не добавляем в список
            If lstMagics.ContainsKey(magicType) = False Then
                Dim lst As New List(Of Integer)
                lst.Add(mgId)
                lstMagics.Add(magicType, lst)
            Else
                lstMagics(magicType).Add(mgId)
            End If
        Next mgId

        'выбираем случайным образом тип используемой магии (исходя из весов)
        Dim mPos As Double = Rnd() * weightsSum, cnt As Double = 0, selectedType As MagicTypesEnum = MagicTypesEnum.Damage
        For i As Integer = 0 To magicWeights.Count - 1
            cnt += magicWeights(i)
            If mPos <= cnt AndAlso magicWeights(i) > 0 Then
                selectedType = i
            End If
        Next i

        'если акая магия одна - выбираем ее
        If lstMagics(selectedType).Count = 1 Then
            GVARS.G_CURMAGIC = lstMagics(selectedType)(0)
            Return 1
        End If

        'если таких магий несколько - выбираем самую дорогую
        Dim arrCost() As Double 'список цен магий
        Dim cost As Double = 0, costMax As Double = -1
        ReDim arrCost(lstMagics(selectedType).Count - 1)

        'создаем список цен магий и получаем самую высокую
        For i As Integer = 0 To lstMagics(selectedType).Count - 1
            Dim mgId As Integer = lstMagics(selectedType)(i)
            If ReadPropertyDbl(classMg, "Cost", magicBookId, mgId, cost, {magicBookId.ToString, mgId.ToString}) = False Then Return -1
            If cost < costMax Then Continue For 'магия заведомо дешевле уже найденной - не рассматриваем
            If costMax < cost Then costMax = cost
            arrCost(i) = cost
        Next

        'создаем список магий, имеющих цену costMax
        Dim lstFinal As New List(Of Integer)
        For i As Integer = 0 To arrCost.Count - 1
            If arrCost(i) = costMax Then lstFinal.Add(lstMagics(selectedType)(i))
        Next
        lstMagics.Clear()
        Erase arrCost

        'из самых дорогих магий выбираем одну случайным образом (если их больше 1)
        If lstFinal.Count = 1 Then
            GVARS.G_CURMAGIC = lstFinal(0)
            Return 1
        End If

        mPos = Math.Round(Rnd() * lstFinal.Count - 0.5)
        If mPos = lstFinal.Count Then mPos -= 1
        GVARS.G_CURMAGIC = lstFinal(CInt(mPos))
        Return 1
    End Function

    ''' <summary>
    ''' Функция удаляет все способности, наложенные умирающим бойцом, свойство RemoveWithApplierDeath которых равно True
    ''' </summary>
    ''' <param name="hostId">Id умирающего бойца</param>
    ''' <returns>"True" если была удалена хоть одна способность</returns>
    Public Function RemoveAbilitiesWithApplierDeath(ByVal hostId As Integer) As String
        Dim classAb As Integer = mScript.mainClassHash("Ab")
        Dim blnWasRemoved As Boolean = False

        For fId As Integer = 0 To Fighters.Count - 1
            'получаем набор способностей
            If Fighters(fId).abilitySetId < 0 Then Continue For
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(Fighters(fId).abilitySetId)("AbilityOnReleaseEvent")
            If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Continue For
            Dim abSetId As Integer = Fighters(fId).abilitySetId

            'с мертвых не снимаем
            Dim life As Double = 0
            If ReadFighterPropertyDbl("Life", fId, life, {fId.ToString}) = False Then Return "#Error"
            If life <= 0 Then Continue For

            For abId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                'перебираем все способности каждого героя
                Dim shouldRemove As Boolean = False
                Dim arrs() As String = {abSetId.ToString, abId.ToString}
                If ReadPropertyBool(classAb, "RemoveWithApplierDeath", abSetId, abId, shouldRemove, arrs) = False Then Return "#Error"
                If Not shouldRemove Then Continue For 'способность не удаляется
                Dim applier As String = "", applierId As Integer = -1
                If ReadProperty(classAb, "Applier", abSetId, abId, applier, arrs) = False Then Return "#Error"
                applierId = GetFighterByName(applier)
                If applierId <> hostId Then Continue For 'хозяин не тот
                'непосредственно удаление
                If FunctionRouter(classAb, "Remove", arrs, Nothing) = "#Error" Then Return "#Error"
                blnWasRemoved = True
            Next abId
        Next fId

        Return "True"
    End Function

    ''' <summary>
    ''' Возвращает Id бойца по его имени (в кавычках)
    ''' </summary>
    ''' <param name="fName">имя в кавычках (или Id)</param>
    Public Function GetFighterByName(ByVal fName As String) As Integer
        If Fighters.Count = 0 Then Return -1
        If IsNumeric(fName) Then
            Dim fighterId As Integer = Val(fName)
            If fighterId < 0 OrElse fighterId > Fighters.Count - 1 Then
                Return -1
            Else
                Return fighterId
            End If
        End If
        For fId As Integer = 0 To Fighters.Count - 1
            If String.Compare(Fighters(fId).heroProps("Name").Value, fName, True) = 0 Then Return fId
        Next fId
        Return -1
    End Function

    ''' <summary>
    ''' Возвращает новое имя бойца на основании шаблона (Крыса2, Крыса3 ...)
    ''' </summary>
    ''' <param name="template">Шаблон</param>
    Public Function GetNewName(ByVal template As String) As String
        If Fighters.Count = 0 Then Return template
        If GetFighterByName(template) < 0 Then Return template

        template = UnWrapString(template)
        Dim i As Integer = 2
        Do
            Dim newName As String = "'" & template & i.ToString & "'"
            If GetFighterByName(newName) < 0 Then Return newName
            i += 1
        Loop
    End Function

    ''' <summary>
    ''' Возвращает количество живых и несбежавших бойцов из указанной армии, учавствующих в бою
    ''' </summary>
    ''' <param name="armyId">Id армии</param>
    Public Function UnitsInArmyLeft(ByVal armyId As Integer) As Integer
        If Fighters.Count = 0 Then Return 0
        Dim unitsLeft As Integer = 0
        For fId As Integer = 0 To Fighters.Count - 1
            Dim fig As cFighters = Fighters(fId)
            If fig.armyId > -1 Then Continue For

            Dim life As Double = 0, runAway As Boolean = False, arrs() As String = {fId.ToString}
            If ReadFighterPropertyDbl("Life", fId, life, arrs) = False Then Return -1
            If ReadFighterPropertyBool("RunAway", fId, runAway, arrs) = False Then Return -1
            If life <= 0 OrElse runAway Then Continue For

            unitsLeft += 1
        Next

        Return unitsLeft
    End Function

    ''' <summary>
    ''' Устанавливает на изображение указанной жертвы эффект повреждения, который определяется исходя из текущей атаки (StrikeType должен быть определен)
    ''' </summary>
    ''' <param name="victimId">Id жертвы</param>
    ''' <param name="hEff">для получения html-элемента (IMG) эффекта</param>
    Public Function SetDamageEffectToVictim(ByVal victimId As Integer, ByRef hEff As HtmlElement) As String
        'эффект удара
        '...
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть html-документ Battle.html!")
        'If GVARS.G_STRIKETYPE < 0 Then Return _ERROR("Не выбран тип удара!")
        If GVARS.G_STRIKETYPE < 0 Then Return ""
        Dim hVictim As HtmlElement = FindFighterHTMLContainer(victimId, hDoc)
        If IsNothing(hVictim) Then Return ""
        hEff = Nothing
        Dim hImg As HtmlElement = FindFighterHTMLImage(hVictim, victimId, hEff)
        If IsNothing(hImg) OrElse IsNothing(hEff) Then Return ""

        Dim effClass As String = "", EffectPicture As String = ""
        Select Case GVARS.G_STRIKETYPE
            Case 0 'обычный удар
                If ReadFighterProperty("EffectClass", victimId, effClass, {victimId.ToString}) = False Then Return "#Error"
                If ReadFighterProperty("EffectPicture", victimId, EffectPicture, {victimId.ToString}) = False Then Return "#Error"
            Case 1 'магия
                If GVARS.G_CURFIGHTER < 0 Then Return _ERROR("Не выбран атакующий боец!")
                Dim mBookId As Integer = Fighters(GVARS.G_CURFIGHTER).magicBookId
                If mBookId < 0 Then Return _ERROR("Магическая атака невозможна. У атакующего персонажа нет книги заклинаний!")
                If GVARS.G_CURMAGIC < 0 Then Return _ERROR("Не выбрана магия для использования!")
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                If ReadProperty(classMg, "EffectClass", mBookId, GVARS.G_CURMAGIC, effClass, {mBookId.ToString, GVARS.G_CURMAGIC.ToString}) = False Then Return "#Error"
                If ReadProperty(classMg, "EffectPicture", mBookId, GVARS.G_CURMAGIC, EffectPicture, {mBookId.ToString, GVARS.G_CURMAGIC.ToString}) = False Then Return "#Error"
            Case Else 'удар другого типа
                Dim propName As String = "EffectClassType" & GVARS.G_STRIKETYPE.ToString
                Dim classH As Integer = mScript.mainClassHash("H")
                If mScript.mainClass(classH).Properties.ContainsKey(propName) = False Then propName = "EffectClass"
                If ReadFighterProperty(propName, victimId, effClass, {victimId.ToString}) = False Then Return "#Error"
                propName = "EffectPictureType" & GVARS.G_STRIKETYPE.ToString
                If mScript.mainClass(classH).Properties.ContainsKey(propName) = False Then propName = "EffectPicture"
                If ReadFighterProperty(propName, victimId, EffectPicture, {victimId.ToString}) = False Then Return "#Error"
        End Select

        effClass = UnWrapString(effClass)
        EffectPicture = UnWrapString(EffectPicture).Replace("\"c, "/"c)
        If String.IsNullOrEmpty(EffectPicture) = False Then
            hEff.Style &= "display:block;"
            hEff.SetAttribute("src", EffectPicture)
            If String.IsNullOrEmpty(effClass) = False Then HTMLAddClass(hEff, effClass)
            Dim msEff As mshtml.HTMLImg = hEff.DomElement
            Do While Not msEff.complete
                Application.DoEvents()
            Loop
        End If

        Return ""
    End Function

#Region "ReadFighterProperty"
    ''' <summary>
    ''' Возвращает обработанне значение свойство бойца (с исполненными скриптами, выбранными селекторами и т. д.)
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="result">для получения результата</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <param name="retFormat">для получения формата свойства</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadFighterProperty(propName As String, ByVal fId As Integer, ByRef result As String, Optional arrParams() As String = Nothing, _
                                 Optional ByRef retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL, Optional selector As Integer = -1) As Boolean
        On Error GoTo er
        retFormat = MatewScript.ReturnFormatEnum.ORIGINAL
        If fId > Fighters.Count - 1 Then
            mScript.LAST_ERROR = String.Format("Бойца с Id = {0} не существует.", fId.ToString)
            Return False
        End If
        If IsNothing(arrParams) Then arrParams = {fId.ToString}

        Dim prop As MatewScript.ChildPropertiesInfoType = Nothing
        Dim props As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Fighters(fId).heroProps
        If props.TryGetValue(propName, prop) = False Then
            mScript.LAST_ERROR = String.Format("Свойство {0} в классе Hero не найдено.", propName)
            Return False
        End If

        Dim value As String, eventId As Integer
        eventId = prop.eventId
        value = prop.Value

        If eventId > 0 Then
            Dim codeType As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(value)
            If codeType = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                'тип - длинный текст

                If selector = -1 Then
                    'селектор не указан
                    'получаем событие селектора
                    Dim propSelector As String = propName & "Selector"
                    'получаем eventId события выбора селектора данного длинного текста (если такое свойство есть)
                    Dim selectorEventId As Integer = 0
                    If props.ContainsKey(propSelector) Then
                        selectorEventId = props(propSelector).eventId
                    End If

                    If selectorEventId > 0 Then
                        'событие выбора селектора описано
                        selector = 0
                        Dim cnt As Integer = GetSelectorsCount(mScript.eventRouter.GetExDataByEventId(eventId)) 'получаем количестово селекторов в тексте свойства
                        If cnt = 0 Then
                            'нет ни одного селектора
                            selector = 0
                        Else
                            Dim selArrs() As String 'параметры для передачи событию выбора селектора
                            If IsNothing(arrParams) OrElse arrParams.Count = 0 Then
                                ReDim selArrs(0)
                                selArrs(0) = cnt.ToString
                            Else
                                ReDim selArrs(arrParams.Count)
                                Array.ConstrainedCopy(arrParams, 0, selArrs, 1, arrParams.Count)
                                selArrs(0) = cnt.ToString
                            End If
                            Dim res As String = mScript.eventRouter.RunEvent(selectorEventId, selArrs, "Hero." & propName & "Selector", False) 'запускаем событие вроде DescriptionSelector, которое возвращает номер селектора для вывода текста
                            If res = "#Error" Then
                                selector = 1 'если произошла ошибка при выборе - берем первое описание (Писатель все-равно увидит сообщение об ошибке и EXIT_CODE все-равно уже True)
                            Else
                                If Integer.TryParse(res, selector) = False Then selector = 1
                            End If
                        End If
                        result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, selector)
                    Else
                        'селектора нет или он незаполнен
                        result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, 0)
                    End If
                Else
                    'указан нужный селектор - получаем описание из нужного селектора
                    result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, selector)
                End If
                retFormat = MatewScript.ReturnFormatEnum.TO_LONG_TEXT
            Else
                'тип - скрипт
                value = mScript.eventRouter.RunEvent(eventId, arrParams, "Свойство Hero." & propName, False)
                If value = "#Error" Then Return False
                result = value
                retFormat = mScript.Param_GetType(value)
            End If
        Else
            'простое значение
            If value.StartsWith("'?") Then
                'executable string
                result = mScript.ExecuteString({value}, arrParams)
                retFormat = MatewScript.ReturnFormatEnum.TO_EXECUTABLE_STRING
            Else
                'теперь точно простое значение
                result = value
                retFormat = mScript.Param_GetType(result)
            End If
        End If

        Return True
er:
        Return False
    End Function

    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате целого числа. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadFighterPropertyInt(ByVal propName As String, ByVal fId As Integer, ByRef result As Integer, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadFighterProperty(propName, fId, strResult, arrParams) = False Then Return False

        If strResult.StartsWith("'"c) Then strResult = UnWrapString(strResult) 'если строка, то получаем ее без кавычек. Это полезно для получения численных Enum
        If IsNumeric(strResult) Then
            result = CInt(strResult)
        Else
            result = Val(strResult)
        End If

        Return True
    End Function

    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате дробного числа. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadFighterPropertyDbl(ByVal propName As String, ByVal fId As Integer, ByRef result As Double, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadFighterProperty(propName, fId, strResult, arrParams) = False Then Return False

        If strResult.StartsWith("'"c) Then strResult = UnWrapString(strResult) 'если строка, то получаем ее без кавычек. Это полезно для получения численных Enum
        If Double.TryParse(strResult, System.Globalization.NumberStyles.Any, provider_points, result) = False Then
            result = Val(strResult)
        End If

        Return True
    End Function

    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате Bool. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="fId">Id бойца</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadFighterPropertyBool(ByVal propName As String, ByVal fId As Integer, ByRef result As Boolean, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadFighterProperty(propName, fId, strResult, arrParams) = False Then Return False

        Boolean.TryParse(strResult, result)
        Return True
    End Function

#End Region

    Private Sub timMoveFighters_Tick(sender As Object, e As EventArgs) Handles timMoveFighters.Tick
        'перемещаем бойцов из-за пределов поля боя на свои места в армиях
        For i As Integer = 0 To lsthFightersToMoveInto.Count - 1
            Dim hFighter As HtmlElement = lsthFightersToMoveInto(i)
            If IsNothing(hFighter) Then Continue For
            hFighter.Style &= "left:0px;top:0px;"
        Next i
        lsthFightersToMoveInto.Clear()

        If lstFightersToMoveOut.Count > 0 Then
            Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
            If IsNothing(hDoc) Then
                timMoveFighters.Enabled = False
                Return
            End If

            'уводим сбежавших бойцов с поля боя
            For i As Integer = 0 To lstFightersToMoveOut.Count - 1
                'перебираем всех сбежавших
                Dim fId As Integer = lstFightersToMoveOut(i)
                Dim runAway As Boolean = False
                ReadFighterPropertyBool("RunAway", fId, runAway, {fId.ToString})
                If Not runAway Then Continue For 'если он уже не сбежавший - пропускаем

                'получаем координаты куда перемещать
                Dim hFighter As HtmlElement = FindFighterHTMLContainer(fId, hDoc)
                If IsNothing(hFighter) Then Continue For
                Dim pt As Point = GetFighterPositionOutOfBattlefield(hFighter)
                'перемещаем
                hFighter.Style &= "left:" & pt.X.ToString & "px;top:" & pt.Y.ToString & "px;"
            Next i

            Dim lst As New List(Of Integer)
            lst.AddRange(lstFightersToMoveOut)
            lstFightersToMoveOut.Clear()
            Wait()

            'а теперь удаляем с поля боя
            For i As Integer = 0 To lst.Count - 1
                'перебираем всех сбежавших
                Dim fId As Integer = lst(i)
                Dim runAway As Boolean = False
                ReadFighterPropertyBool("RunAway", fId, runAway, {fId.ToString})
                If Not runAway Then Continue For 'если он уже не сбежавший - пропускаем

                'удаляем
                RemoveFighterFromBattlefield(fId)
            Next i

        End If

        timMoveFighters.Enabled = False
    End Sub
End Class
