Public Class CalcStats

    Enum FormulaType As Integer
        Fallout2 = 0
        Fallout1 = 1
        Custom = 2
    End Enum

    Private Shared formula As FormulaType = FormulaType.Fallout2

    Friend Shared Sub SetFormula(ByVal type As FormulaType)
        formula = type
    End Sub

    Friend Shared Function GetFormula() As Integer
        Return formula
    End Function

    Friend Shared Function SmallGun_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (35 + (1 * Agility)), (5 + (4 * Agility)))
    End Function

    Friend Shared Function BigEnergyGun_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (10 + (1 * Agility)), (2 * Agility))
    End Function

    Friend Shared Function EnergyGun_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (10 + (1 * Agility)), (2 * Agility))
    End Function

    Friend Shared Function Unarmed_Skill(ByVal Agility As Integer, ByVal Strength As Integer) As Integer
        Return CInt(If(CBool(formula), (65 + (1 * (Agility + Strength) / 2)), (30 + (Agility + Strength) * 2)))
    End Function

    Friend Shared Function Melee_Skill(ByVal Agility As Integer, ByVal Strength As Integer) As Integer
        Return CInt(If(CBool(formula), (55 + (1 * (Agility + Strength) / 2)), (20 + (Agility + Strength) * 2)))
    End Function

    Friend Shared Function Throwing_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (40 + (1 * Agility)), (4 * Agility))
    End Function

    Friend Shared Function Action_Point(ByVal Agility As Integer) As Integer
        Return CInt(Fix(5 + (Agility / 2)))
    End Function

    Friend Shared Function Health_Point(ByVal Strength As Integer, ByVal Endurance As Integer) As Integer
        Return (15 + Strength + (Endurance * 2))
    End Function

    Friend Shared Function Healing_Rate(ByVal Endurance As Integer) As Integer
        Return CInt(Math.Max(1, Int((Endurance / 3))))
    End Function

    Friend Shared Function Melee_Damage(ByVal Strength As Integer) As Integer
        Return Math.Max(1, (Strength - 5))
    End Function

    Friend Shared Function Sequence(ByVal Perception As Integer) As Integer
        Return (2 * Perception)
    End Function

    Friend Shared Function Radiation(ByVal Endurance As Integer) As Integer
        Return (2 * Endurance)
    End Function

    Friend Shared Function Poison(ByVal Endurance As Integer) As Integer
        Return (5 * Endurance)
    End Function
    '
    Friend Shared Function Carry_Weight(ByVal Strength As Integer) As Integer
        Return (25 + (25 * Strength))
    End Function

    Friend Shared Function FirstAid_Skill(ByVal Perception As Integer, ByVal Intelligence As Integer) As Integer
        Return CInt(If(CBool(formula), (30 + (1 * (Perception + Intelligence) / 2)), ((Perception + Intelligence) * 2)))
    End Function

    Friend Shared Function Doctor_Skill(ByVal Perception As Integer, ByVal Intelligence As Integer) As Integer
        Return CInt(If(CBool(formula), (15 + (1 * (Perception + Intelligence) / 2)), (5 + Perception + Intelligence)))
    End Function

    Friend Shared Function Outdoorsman_Skill(ByVal Endurance As Integer, ByVal Intelligence As Integer) As Integer
        Return CInt(If(CBool(formula), (5 + (1 * (Endurance + Intelligence) / 2)), ((Endurance + Intelligence) * 2)))
    End Function

    Friend Shared Function Sneak_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (25 + (1 * Agility)), (5 + (3 * Agility)))
    End Function

    Friend Shared Function Lockpick_Skill(ByVal Perception As Integer, ByVal Agility As Integer) As Integer
        Return CInt(If(CBool(formula), (20 + (1 * (Perception + Agility) / 2)), (10 + Perception + Agility)))
    End Function

    Friend Shared Function Steal_Skill(ByVal Agility As Integer) As Integer
        Return If(CBool(formula), (20 + (1 * Agility)), (3 * Agility))
    End Function

    Friend Shared Function Trap_Skill(ByVal Perception As Integer, ByVal Agility As Integer) As Integer
        Return CInt(If(CBool(formula), (20 + (1 * (Perception + Agility) / 2)), (10 + Perception + Agility)))
    End Function

    Friend Shared Function Science_Skill(ByVal Intelligence As Integer) As Integer
        Return If(CBool(formula), (25 + (2 * Intelligence)), (4 * Intelligence))
    End Function

    Friend Shared Function Repair_Skill(ByVal Intelligence As Integer) As Integer
        Return If(CBool(formula), (20 + (1 * Intelligence)), (3 * Intelligence))
    End Function

    Friend Shared Function Speech_Skill(ByVal Charisma As Integer) As Integer
        Return If(CBool(formula), (25 + (2 * Charisma)), (5 * Charisma))
    End Function

    Friend Shared Function Barter_Skill(ByVal Charisma As Integer) As Integer
        Return If(CBool(formula), (20 + (2 * Charisma)), (4 * Charisma))
    End Function

    Friend Shared Function Gamblings_Skill(ByVal Luck As Integer) As Integer
        Return If(CBool(formula), (20 + (3 * Luck)), (5 * Luck))
    End Function

End Class
