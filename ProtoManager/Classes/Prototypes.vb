﻿Public Class Prototypes

    '
    ' Описание структур для файлов прототипов
    '

    'Cтруктура для файла типа Critters
    Structure CritPro
        'Comon
        Public ProtoID As Integer
        Public DescID As Integer
        Public FrmID As Integer
        Public LightDis As Integer
        Public LightInt As Integer
        'Critter
        Public Falgs As Integer
        Public FalgsExt As Integer
        Public ScriptID As Integer
        Public HeadFID As Integer
        Public AIPacket As Integer
        Public TeamNum As Integer
        Public CritterFlags As Integer
        'Base param
        Public Strength As Integer
        Public Perception As Integer
        Public Endurance As Integer
        Public Charisma As Integer
        Public Intelligence As Integer
        Public Agility As Integer
        Public Luck As Integer
        Public HP As Integer
        Public AP As Integer
        Public AC As Integer
        Public UnarmedDmg As Integer
        Public MeleeDmg As Integer
        Public Weight As Integer
        Public Sequence As Integer
        Public Healing As Integer
        Public Critical As Integer
        Public Better As Integer
        'Base Protect
        Public DTNormal As Integer
        Public DTLaser As Integer
        Public DTFire As Integer
        Public DTPlasma As Integer
        Public DTElectrical As Integer
        Public DTEMP As Integer
        Public DTExplode As Integer
        Public DRNormal As Integer
        Public DRLaser As Integer
        Public DRFire As Integer
        Public DRPlasma As Integer
        Public DRElectrical As Integer
        Public DREMP As Integer
        Public DRExplode As Integer
        Public DRRadiation As Integer
        Public DRPoison As Integer
        '
        Public Age As Integer
        Public Gender As Integer
        'Bonus to base param
        Public b_Srength As Integer
        Public b_Perception As Integer
        Public b_Endurance As Integer
        Public b_Charisma As Integer
        Public b_Intelligence As Integer
        Public b_Agility As Integer
        Public b_Luck As Integer
        Public b_HP As Integer
        Public b_AP As Integer
        Public b_AC As Integer
        Public b_UnarmedDmg As Integer
        Public b_MeleeDmg As Integer
        Public b_Weight As Integer
        Public b_Sequence As Integer
        Public b_Healing As Integer
        Public b_Critical As Integer
        Public b_Better As Integer
        'Bonus Protect
        Public b_DTNormal As Integer
        Public b_DTLaser As Integer
        Public b_DTFire As Integer
        Public b_DTPlasma As Integer
        Public b_DTElectrical As Integer
        Public b_DTEMP As Integer
        Public b_DTExplode As Integer
        Public b_DRNormal As Integer
        Public b_DRLaser As Integer
        Public b_DRFire As Integer
        Public b_DRPlasma As Integer
        Public b_DRElectrical As Integer
        Public b_DREMP As Integer
        Public b_DRExplode As Integer
        Public b_DRRadiation As Integer
        Public b_DRPoison As Integer
        '
        Public b_Age As Integer
        Public b_Gender As Integer
        'Skills
        Public SmallGuns As Integer
        Public BigGuns As Integer
        Public EnergyGun As Integer
        Public Unarmed As Integer
        Public Melee As Integer
        Public Throwing As Integer
        Public FirstAid As Integer
        Public Doctor As Integer
        Public Sneak As Integer
        Public Lockpick As Integer
        Public Steal As Integer
        Public Traps As Integer
        Public Science As Integer
        Public Repair As Integer
        Public Speech As Integer
        Public Barter As Integer
        Public Gambling As Integer
        Public Outdoorsman As Integer
        '
        Public BodyType As Integer
        Public ExpVal As Integer
        Public KillType As Integer
        Public DamageType As Integer
    End Structure

    'Cтруктура для файла типа Items
    'Common
    Structure CmItemPro
        Public ProtoID As Integer
        Public DescID As Integer
        Public FrmID As Integer
        Public LightDis As Integer
        Public LightInt As Integer
        Public Falgs As Integer
        Public FalgsExt As Integer
        Public ScriptID As Integer
        Public ObjType As Integer
        Public MaterialID As Integer
        Public Size As Integer
        Public Weight As Integer
        Public Cost As Integer
        Public InvFID As Integer
        Public SoundID As Byte
    End Structure
    'Weapon
    Structure WpItemPro
        Public AnimCode As Integer
        Public MinDmg As Integer
        Public MaxDmg As Integer
        Public DmgType As Integer
        Public MaxRangeP As Integer
        Public MaxRangeS As Integer
        Public ProjPID As Integer
        Public MinST As Integer
        Public MPCostP As Integer
        Public MPCostS As Integer
        Public CritFail As Integer
        Public Perk As Integer
        Public Rounds As Integer
        Public Caliber As Integer
        Public AmmoPID As Integer
        Public MaxAmmo As Integer
        Public wSoundID As Byte
    End Structure
    'Armor
    Structure ArItemPro
        Public AC As Integer
        Public DRNormal As Integer
        Public DRLaser As Integer
        Public DRFire As Integer
        Public DRPlasma As Integer
        Public DRElectrical As Integer
        Public DREMP As Integer
        Public DRExplode As Integer
        Public DTNormal As Integer
        Public DTLaser As Integer
        Public DTFire As Integer
        Public DTPlasma As Integer
        Public DTElectrical As Integer
        Public DTEMP As Integer
        Public DTExplode As Integer
        Public Perk As Integer
        Public MaleFID As Integer
        Public FemaleFID As Integer
    End Structure
    'Ammo
    Structure AmItemPro
        Public Caliber As Integer
        Public Quantity As Integer
        Public ACAdjust As Integer
        Public DRAdjust As Integer
        Public DamMult As Integer
        Public DamDiv As Integer
    End Structure
    'Key
    Structure kItemPro
        Public Unknown As Integer
    End Structure
    'Misc
    Structure McItemPro
        Public PowerPID As Integer
        Public PowerType As Integer
        Public Charges As Integer
    End Structure
    'Container
    Structure CnItemPro
        Public MaxSize As Integer
        Public OpenFlags As Integer
    End Structure
    'Drugs
    Structure DgItemPro
        Public Stat0 As Integer
        Public Stat1 As Integer
        Public Stat2 As Integer
        Public iAmount0 As Integer
        Public iAmount1 As Integer
        Public iAmount2 As Integer
        Public Duration1 As Integer
        Public fAmount0 As Integer
        Public fAmount1 As Integer
        Public fAmount2 As Integer
        Public Duration2 As Integer
        Public sAmount0 As Integer
        Public sAmount1 As Integer
        Public sAmount2 As Integer
        Public AddictionRate As Integer
        Public W_Effect As Integer
        Public W_Onset As Integer
    End Structure

End Class
