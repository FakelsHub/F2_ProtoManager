Public Class CalcStats

    Friend Function SmallGun_Skill() As Integer
        Return (5 + (4 * ReverseBytes(CritterPro.Agility)))
    End Function

    Friend Function BigEnergyGun_Skill() As Integer
        Return (2 * ReverseBytes(CritterPro.Agility))
    End Function

    'Friend Function EnergyGun_Skill(ByRef Agility) As Integer
    '    Return (2 * ReverseBytes(Agility))
    'End Function

    Friend Function Melee_Skill() As Integer
        Return (20 + (ReverseBytes(CritterPro.Agility) + ReverseBytes(CritterPro.Strength)) * 2)
    End Function

    Friend Function Unarmed_Skill() As Integer
        Return (30 + (ReverseBytes(CritterPro.Agility) + ReverseBytes(CritterPro.Strength)) * 2)
    End Function

    Friend Function Throwing_Skill() As Integer
        Return (4 * ReverseBytes(CritterPro.Agility))
    End Function

    Friend Function Action_Point() As Integer
        Return Fix(5 + (ReverseBytes(CritterPro.Agility) / 2))
    End Function

    Friend Function Health_Point() As Integer
        Return (15 + ReverseBytes(CritterPro.Strength) + (ReverseBytes(CritterPro.Endurance) * 2))
    End Function

    Friend Function Healing_Rate() As Integer
        Return Math.Max(1, Int((ReverseBytes(CritterPro.Endurance) / 3)))
    End Function

    Friend Function Melee_Damage() As Integer
        Return Math.Max(1, (ReverseBytes(CritterPro.Strength) - 5))
    End Function

    Friend Function Sequence() As Integer
        Return (2 * ReverseBytes(CritterPro.Perception))
    End Function

    Friend Function Radiation() As Integer
        Return (2 * ReverseBytes(CritterPro.Endurance))
    End Function

    Friend Function Poison() As Integer
        Return (5 * ReverseBytes(CritterPro.Endurance))
    End Function

End Class
