Attribute VB_Name = "TURN ACTIVITIES"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                          MAINTENANCE LOG                              *****'
'*                              VERSION 3.2.3                                      *'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 05/02/96  *  REMOVE MONEY CALCS                                             **'
'** 19/02/96  *  CHANGE THE ACTIVITY TABLE BEING ACCESSED.                      **'
'**           *  FIX WHALES & MINING PROBLEMS                                   **'
'** 23/02/96  *  fix provs eaten, allow for appropriate mining tool & companion **'
'**               *  metal                                                      **'
'** 25/02/96  *  Plowing & Planting are now reported.                           **'
'**               *  Charcoal production is now reported.                       **'
'** 24/07/96  *  Allow for Cooking, Music, Art, Dancing                         **'
'** 16/08/96  *  Allow for Writing Books                                        **'
'** 18/10/96  *  Allow for Camels                                               **'
'** 13/11/97  *  Allow for Hemp                                                 **'
'** 17/08/99  *  Fix prolem with Low Yield Mining                               **'
'** 04/10/99  *  Add in code for seeking skill levels.                          **'
'** 31/10/01  *  REBUILD TO CATER FOR TABLE INPUT                               **'
'**                                                                             **'
'** 13/11/01 - module is currently 169 pages in length                          **'
'** 14/11/01 - MODULE IS CURRENTLY 160 PAGES IN LENGTH - CONSTRUCTION MODULE HAS BEEN DELETED
'** 20/03/18  *  Adjustments to activity printing, add fish to final activities **'
'** 06/04/20  * Changed Breed New Queens to just give 24 hives every Spring01   **'
'** 08/04/20  * Placed a check in place to ensure if an incorrect activity is   **'
'**           * entered then it will be ignored and processing will continue    **'
'**                                                                             **'
'** 05/03/25  * Apotecary implementation added (AlexD)                          **'
'** 05/03/25  *                                           **'
'** 05/03/25  * Salting added to Fishing, so that unused fish may be salted. (AlexD) **'
'** 05/03/25  * Engineering of container/Installation buildings rewritten (AlexD)    **'
'** 05/03/25  * Added eligibility checks for engineering  (AlexD)               **'
'** 05/03/25  * Farming of non-permanent crops fixed (AlexD)                    **'
'** 05/03/25  * Fixing wrong resource consumption for 4 skills:                 **'
'**                BONEWORK, CURING, DRESSING and STONEWORK (AlexD)             **'
'** 05/03/25  * Fixing wrong food consumption via GT (fish, milk and bread) (AlexD)  **'
'** 05/03/25  * Fixing Iron Burner problem     (AlexD)                          **'
'** 05/03/25  * Enabled food consumption print (AlexD)                          **'
'** 19/03/25  * Disabled "TOWER WOODEN" (AlexD)                                 **'
'** 20/03/25  * Fixed farming in adjacient hex  (AlexD)                         **'
'** 20/03/25  * Fixed fish dissapearance (AlexD)                                **'
'** 20/03/25  * Hirelings Added (AlexD)                                         **'
'** 28/03/25  * Bread consumption fixed (AlexD)                                 **'
'** 01/04/25  * Removed requirement that adjacent MH should belong to GT for farming (AlexD)                                 **'
'*===============================================================================*'

' MODULE NAME IS TURN ACTIVITIES
' Number variables.  Long has no decimal points, double has decimal points
Global Mycontrol As Control
Global TVDB As DAO.Database
Global TVDBGM As DAO.Database
Global TVWKSPACE As Workspace

'******  TABLES  ********
Global activitytable As Recordset
Global ActivitiesTable As Recordset
Global AVAIL_RES_TABLE As Recordset
Global ItemsTable As Recordset
Global ACTSEQTAB As Recordset
Global COSTSTABLE As Recordset
Global GAMES_WEATHER As Recordset
Global globalinfo As Recordset
Global GMTABLE As Recordset
Global HEXMAPCITY As Recordset
Global HEXMAPCONST As Recordset
Global HEXMAPMINERALS As Recordset
Global HEXMAPPOLITICS As Recordset
Global hexmaptable As Recordset
Global HSTABLE As Recordset
Global ImplementsTable As Recordset
Global ImplementUsage As Recordset
Global OutPutTable As Recordset
Global PopTable As Recordset
Global SEASONTABLE As Recordset
Global SKILLSTABLE As Recordset
Global FarmingTable As Recordset
Global CROP_TABLE As Recordset
Global ConstructionTable As Recordset
Global Building_Used As Recordset
Global COMPRESTAB As Recordset
Global MODTABLE As Recordset
Global CLIMATETABLE As Recordset
Global LINENUMTAB As Recordset
Global Goods_Tribes_Processed As Recordset
Global PermFarmingTable As Recordset
Global PACIFICATION_TABLE As Recordset
Global PROVS_AVAIL_TABLE As Recordset
Global PROCESSACTIVITY As Recordset
Global PROCESSITEMS As Recordset
Global RESEARCH_TABLE As Recordset
Global MOVEMENT_TABLE As Recordset
Global SCOUTING_TABLE As Recordset
Global SeekingReturnsTable As Recordset
Global SHIPSTABLE As Recordset
Global TRIBES_CHECKING As Recordset
Global TRIBESTABLE As Recordset
Global TRIBESGOODS As Recordset
Global TribesGoodsUsage As Recordset
Global TRIBESINFO As Recordset
Global TribesModifiers As Recordset
Global TrAct_Req_Later As Recordset
Global Turn_Info_Req_NxTurn As Recordset
Global Turns_Activities_Table As Recordset
Global TERRAINTABLE As Recordset
Global TEMPCOMPRESTAB As Recordset
Global TRIBESBOOKS As Recordset
Global TRADING_POST_GOODS As Recordset
Global TRIBES_ACTIVITY_IMPLEMENTS As Recordset
Global Tribes_Processing As Recordset
Global TRIBE_RESEARCH As Recordset
Global TribesSpecialists As Recordset
Global VALID_CONST As Recordset
Global VALID_DIRECTIONS As Recordset
Global VALIDANIMALS As Recordset
Global VALIDMINERALS As Recordset
Global VALIDSHIPS As Recordset
Global WEATHERTABLE As Recordset
Global WINDTABLE As Recordset
Global Import_Trades As Recordset
Global ConstructionTable2 As Recordset
Global MASS_TRANSFER_COPY_TABLE As Recordset
Global Process_Tribe_Movement_Copy As Recordset
Global Process_Skills_Copy As Recordset
Global Process_Research_Copy As Recordset
Public Printing_Switch_TABLE As Recordset
Global SCOUT_MOVEMENT_TABLE As Recordset
Global SCOUT_MOVEMENT_COPY As Recordset
Global TRIBE_MOVEMENT_TABLE As Recordset
Global Movement_Trace As Recordset
Global Scout_Result As Recordset


'******  A100 VARIABLES  ********
Global CURRENT_ACTIVITY As String
Global TCLANNUMBER As String
Global TTRIBENUMBER As String
Global SKILLSTRIBENUMBER As String
Global TActivity As String
Global TDistinction As String
Global TItem As String
Global TJoint As String
Global TActives As Long
Global TempActives As Long
Global TInActives As Long
Global TSlaves As Long
Global TSpecialists As Long
Global TOwning_Clan As String
Global TOwning_Tribe As String
Global Number_of_Seeking_Groups As Long
Global Number_of_Seeking_Attempts As Long
Global TWhale_Size As String
Global TMining_Direction As String
Global INVALID_TRIBE As String
Global TBuilding As Long
Global Output_Processed As String

'******  A150 ACTIVITY VARIABLES  ********
Global No_Activity_Record_Found As String
Global TPeople As Long
Global TResearch As String
Global TShort As String
Global TItem_Produced As String
Global TSkilllvl(4) As Long
Global TSkill(4) As String
Global TNumItems As Long
Global TBuildingLimit As Long
Global TManufacturingLimit As Long
Global TGoods(40) As String
Global TQuantity(40) As Long
Global SQuantity(40) As Long

'******  A250 TRIBES VARIABLES  ********
Global TActivesAvailable As Long
Global TSlavesAvailable As Long
Global TMouths As Long
Global NUMBER_OF_SLAVES As Long
Global GOODS_CLAN As String
Global GOODS_TRIBE As String
Global Tribes_Current_Hex As String
Global CONST_Tribes_Current_Hex As String
Global Goods_Tribes_Current_Hex As String
Global Meeting_House_Hex As String
Global TRIBES_TERRAIN As String
Global GOODS_TRIBES_TERRAIN As String
Global Government_Level As Long
Global Skill_Tribe As String
Global CURRENT_WEATHER_ZONE As String
Global TRIBES_RELIGION As String
Global TRIBES_CULT As String
Global TRIBES_POP_TRIBE As String   ' ALSO USED IN PROCESS_POPULATION_GROWTH
Global TRIBES_COST_CLAN As String
Global TRIBES_MORALE As Double

'******  A250 HEX VARIABLES  ********
Global FARMING_TERRAIN As String
Global FRESH_WATER As Single
Global ROAMING_HERD As String
Global QUARRYING As String
Global SALMON_RUN As String
Global FISH_AREA As String
Global WHALE_AREA As String
Global RIVER_N As String
Global RIVER_NE As String
Global RIVER_SE As String
Global RIVER_S As String
Global RIVER_SW As String
Global RIVER_NW As String
Global Hexmaps_WEATHER_ZONE As String
Global CURRENT_HEX_POP As Long
Global CURRENT_HEX_PAC_LEV As Long

'******  A250 WEATHER VARIABLES  ********
Global CURRENT_WEATHER As String
Global CURRENT_WIND As String
Global CURRENT_CLIMATE As String
Global SEASON_FISHING As Long
Global WEATHER_HERDING_GROUP_1 As Single
Global WEATHER_HERDING_GROUP_2 As Single
Global WEATHER_HERDING_GROUP_3 As Single
Global HUNTING_WEATHER As Single
Global MINING_WEATHER As Single
Global MINING_ACCIDENTS_WEATHER As Single
Global HONEY_WEATHER As Single
Global WAX_WEATHER As Single
Global FISHING_WEATHER As Single

'******  A250 TRIBES SKILLS VARIABLES  ********
Global APIARISM_LEVEL As Long
Global ARCHERY_LEVEL As Long
Global ARMOUR_LEVEL As Long
Global BONING_LEVEL As Long
Global COMBAT_LEVEL As Long
Global DIPLOMACY_LEVEL As Long
Global FARMING_LEVEL As Long
Global FISHING_LEVEL As Long
Global FLENSING_LEVEL As Long
Global FORESTRY_LEVEL As Long
Global FURRIER_LEVEL As Long
Global GUTTING_LEVEL As Long
Global HEALING_LEVEL As Long
Global HERDING_LEVEL As Long
Global HORSEMANSHIP_LEVEL As Long
Global HUNTING_LEVEL As Long
Global HVYWPNS_LEVEL As Long
Global LEADERSHIP_LEVEL As Long
Global MINING_LEVEL As Long
Global NAVIGATION_LEVEL As Long
Global PEELING_LEVEL As Long
Global POLITICS_LEVEL As Long
Global RELIGION_LEVEL As Long
Global ROWING_LEVEL As Long
Global SAILING_LEVEL As Long
Global SALTING_LEVEL As Long
Global SANITATION_LEVEL As Long
Global SCOUTING_LEVEL As Long
Global SEAMANSHIP_LEVEL As Long
Global SECURITY_LEVEL As Long
Global SEEKING_LEVEL As Long
Global SKINNING_LEVEL As Long
Global SLAVERY_LEVEL As Long
Global SPYING_LEVEL As Long
Global TACTICS_LEVEL As Long
Global TORTURE_LEVEL As Long
Global WHALING_LEVEL As Long

'******  A250 TERRAIN VARIABLES  ********
Global TERRAIN_HERDING_GROUP_1 As Single
Global TERRAIN_HERDING_GROUP_2 As Single
Global TERRAIN_HERDING_GROUP_3 As Single
Global TERRAIN_HUNTING As Single

'******  A250 WIND VARIABLES  ********
Global WIND_FISHING As Single

'******  A250 SEASON VARIABLES  ********
Global CURRENT_SEASON As String
Global SEASON_HONEY As Single
Global SEASON_WAX As Single
Global COASTAL_FISHING As Single
Global OCEAN_FISHING As Single
Global FURRIER_SKINS As Single
Global FURRIER_FURS As Single

'******  A250 skill check VARIABLES  ********
Global Allskillsok As String
Global TSkillok(4) As String
Global Skill_Level_1 As Long
Global Skill_Level_2 As Long
Global Skill_Level_3 As Long
Global SKILL_LEVEL_4 As Long
Global SKILL_SHORTAGE As Long
Global MAXIMUM_ACTIVES_1 As Long
Global MAXIMUM_ACTIVES_2 As Long
Global MAXIMUM_ACTIVES_3 As Long
Global MAXIMUM_ACTIVES_4 As Long

'*** APIARISM VARIABLES ***
Global TNewWax As Long, TNewHoney As Long
Global TNewRoyalJelly As Long, TNewHives As Long
Global TNewPropolis As Long

'*** BLUBBERWORK VARIABLES ***
Global AMOUNT_OF_BLUBBER As Long
Global TNewOil As Long

' *** FARMING VARIABLES ***

Global CROP_TYPE As String
Global ACRES_PLANTED As Long
Global ACRES_HARVESTED As Long
Global ACRES_NOT_PLANTED As Long
Global ACRES_TO_PLANT As Long
Global ACRES_TO_MAINTAIN As Long
Global ACRES_MAINTAINED As Long
Global ACRES_PLOWED As Long
Global PLANTING_STARTED As String
Global NEW_FLAX As Long
Global NEW_GRAIN As Long
Global NEW_GRAPES As Long
Global NEW_SUGAR As Long
Global NEW_COTTON As Long
Global NEW_LINSEED As Long
Global NEW_TOBACCO As Long
Global NEW_HEMP As Long
Global NEW_POTATOES As Long
Global TOTAL_FLAX As Long
Global TOTAL_GRAIN As Long
Global TOTAL_LINSEED As Long
Global TOTAL_COTTON As Long
Global TOTAL_SUGAR As Long
Global TOTAL_GRAPES As Long
Global TOTAL_TOBACCO As Long
Global TOTAL_HEMP As Long
Global TOTAL_POTATOES As Long
Global FARMERS As Long
Global WEATHER_TURN1 As String
Global WEATHER_TURN2 As String
Global WEATHER_TURN3 As String
Global TScythes As Long
Global NEW_CROP As Long
Global TOTAL_CROP As Long
Global WEATHER_CROP As String
Global FARMING_TURN1 As String, FARMING_TURN2 As String, FARMING_TURN3 As String
Global HARVEST_CONTINUE As String

' *** HERDING VARIABLES ***

Global TNewGoats As Long, TempGoats As Single
Global TNewCattle As Long, TempCattle As Single
Global TNewCamels As Long, TempCamels As Single
Global TNewHorses As Long, TNewLHorses As Long, TNewHHorses As Long, TempHorses As Single
Global TNewSheep As Long, TempSheep As Single
Global TNewDogs As Long, TempDogs As Single
Global TNewElephants As Long, TempElephants As Single


'*** WHALING VARIABLES ***
Global Num_Whales As Long


'*** POLITICS VARIABLES ***
Global Starting_Hex_GL0 As String


'*** CHECK FOR BUILDINGS VARIABLES ***
Global MEETING_HOUSE_FOUND As String
Global CONST_MEETING_HOUSE_FOUND As String
Global BUILDING_FOUND As String
Global RESEARCH_FOUND As String
Global APIARYS_FOUND As Long
Global HOSPITAL_FOUND As Long
Global SEWER_FOUND As Long
Global MIDWIFERY_FOUND As String



'*** CHECK FOR SPECIALISTS VARIABLES ***
Global SPECIALIST_FOUND As String
Global BAKER_FOUND As Long
Global FORESTER_FOUND As Long
Global HUNTER_FOUND As Long
Global FARMER_FOUND As Long
Global NO_SPECIALISTS_FOUND As Long


'******  COMMON VARIABLES  ********
Global QUERY_STRING As String
Global CURRENT_TERRAIN As String
Global LINEFEED As String
Global FILETV As String
Global FILEGM As String
Global TProvisions As Long
Global TempProvs As Single
Global ACTIVES_INUSE As Long
Global Groups As Long
Global TFishing As Long, Firstmodify As String
Global TFish_Caught As Long
Global TNEWORE As Long
Global Whales As String, NumWhales As Long
Global FODDER_GATHER As Long
Global TEngTribe As String
Global Skill As String
Global LIQUID_STORAGE As Long
Global LIQUID_ONHAND As Long
Global LIQUID_STORAGE_AVAILABLE As Long
Global TNUMOCCURS As Long
Global Index1 As Long, Index2 As Long
Global Index3 As Long, Index4 As Long
Global TCount As Long, Current_Increase As Long
Global NumItemsMade As Long, LINENUMBER As Long
Global TurnActOutPut As String, UpdateGoods As String
Global UPDATEITEMS As String, ModifyTable As String
Global TempNumOccurs As Single
Global TEMPITEM As Long, PercentComplete As Single
Global TOldBuilding As String, TNewBuilding As String
Global HERDERSREQ As Long
Global TACLAN As String
Global TATRIBE As String
Global TAACTIVITY As String
Global TAITEM As String
Global TADISTINCTION As String
Global TAACTIVES As Long
Global CLIMATE As String
Global Current_Turn As String
Global Animals_On_Hand As Long
Global JOINT_TRIBE As String
Global NO_UPDATE_POP As String
Global EXTRA_ACTIVITIES As String
Global qdfCurrent As QueryDef
Global CURRENT_DIRECTORY As String
Global MAIN_GROUP As String
Global MAIN_INCREASE As Long
Global BOOKTOPIC As String
Global ITEM As String
Global MORALELOSS As String
Global TURN_NUMBER As String
Global Msg As String
Global MSG0 As String
Global MSG1 As String
Global MSG2 As String
Global MSG3 As String
Global MSG4 As String
Global MSG5 As String
Global MSG6 As String
Global MSG7 As String
Global MSG8 As String
Global total_available As Long
Global String_Length As Long
Global TRIBE_STATUS As String
Global EXECUTION_STATUS As String
Global IMPLEMENT As String
Global WEATHER As String
Global TOTAL_POPULATION As Long
Global CURRENT_HEX As String
Global AvailableSerfs As Long
Global MAP_N As String
Global MAP_NE As String
Global MAP_SE As String
Global MAP_S As String
Global MAP_SW As String
Global MAP_NW As String
Global STONES_QUARRIED As Long
Global TOTAL_SAND As Long
Global Whale As String
Global TOTAL_CLAY As Long
Global WARHORSES As Long
Global BRICKS_CREATED As Long
Global GRAINREQ As Long
Global FODDERREQ As Long
Global EXTRA_PROVS As Long
Global TOTAL_PROVS_AVAILABLE As Long
Global TProvsReq As Long
Global VILLAGE_FOUND As String
Global ActivesNeeded As Long
Global BRACKET As Long
Global Finished As String
Global DICE_TRIBE As Long
Global TLogs As Long, TBark As Long
Global THeads As Long
Global Silver_Tribute As Long
Global TImplement As Long
Global THunters As Long
Global TQuarriers As Long
Global Initial_Metalworkers As Long
Global TMetalworkers As Long
Global Initial_Armourers As Long
Global TArmourers As Long
Global TMiners As Long
Global TLossMiners As Long
Global TAccident As Long
Global SlaveIncrease As Long
Global GoatsKilled As Long, CattleKilled As Long, HorsesKilled As Long
Global DogsKilled As Long, CamelsKilled As Long
Global TFINISHED As String
Global SLAVES_OVERSEEN As String
Global roll1 As Long
Global roll2 As Long
Global roll3 As Long
Global QUESTION As String
Global FLENSERS As Long
Global PEELERS As Long
Global BONERS As Long
Global NO_FARMING As String
Global TShovels As Long
Global SWAPSINEFFECT As Long
Global TOTALMILK As Long
Global TActivesAtStart As Long
Global TDOGS As Long
Global THORSES As Long
Global ENGINEERS_AVAILABLE As Long
Global QUARRYING_ACTIVES As Long
Global SHIPBUILDING_ACTIVES As Long
Global IMPLEMENT_MODIFIER As Double
Global FIND_CHANCE As Long
Global SEASON_FACTOR As Long
Global NUMBER_OF_RECRUITS As Long
Global POPULATION_STARVED As Long
Global FORESTRY_OK As String
Global MORALE As Double
Global Num_Goods As Long
Global WATER_ONHAND As Long
Global stext As String
Global sValue As String
Global sValue1 As String
Global sValue2 As String
Global PREV_CLANNUMBER As String
Global Number_Of_Implements As Integer
Global Number_Found As Long
Global TempOutput As String

' *** MODIFIERS VARIABLES ***
Global TRAPS_TO_USE As Long
Global SNARES_TO_USE As Long
Global STONES_TO_QUARRY As Long
Global STONES_TO_USE As Long
Global BARK_TO_STRIP As Long
Global LOGS_TO_CUT As Long
Global SCRAPERS_TO_USE As Long
Global Slave_Population_Increase As Double
Global Tribe_Population_Increase As Double

'*** FINAL ACTIVITIES VARIABLES ***
Global HEX_POP_GROWTH As Long
Global TOldWarriors As Long
Global TOldActives As Long
Global TOldInactives As Long
Global TOldPopFigure As Long
Global Ring1(6) As String
Global Ring2(12) As String
Global Ring3(18) As String
Global Ring4(24) As String
Global Ring5(30) As String
Global Ring6(36) As String
Global MOVEMENT(6) As String
Global Iteration_Count(6) As Integer

'*** CONSTRUCTION VARIABLES ****
Global Job As String ' JOB IS ANOTHER NAME FOR THE ACTIVITY IN ACTIVITY TABLE
Global CONSTRUCTION As String  ' IS ANOTHER NAME FOR ITEM
Global CONSTRUCTION_TYPE As String
Global BUILDING_TYPE As String
Global WORKERS As Long   ' IS ANOTHER NAME FOR ACTIVES
Global WORKCLAN As String
Global WORKTRIBE As String
Global CONSTCLAN As String
Global CONSTTRIBE As String
Global GOODSTRIBE As String
Global TOTAL_LOGS As Long
Global TOTAL_HLOGS As Long
Global TOTAL_STONES As Long
Global TOTAL_COAL As Long
Global TOTAL_BRASS As Long
Global TOTAL_BRONZE As Long
Global TOTAL_COPPER As Long
Global TOTAL_IRON As Long
Global TOTAL_LEAD As Long
Global TOTAL_CLOTH As Long
Global TOTAL_LEATHER As Long
Global TOTAL_ROPES As Long
Global TOTAL_MILLSTONES As Long
Global LOGSUSED As Long
Global HLOGSUSED As Long
Global STONESUSED As Long
Global COALUSED As Long
Global BRASSUSED As Long
Global BRONZEUSED As Long
Global COPPERUSED As Long
Global IRONUSED As Long
Global LEADUSED As Long
Global CLOTHUSED As Long
Global LEATHERUSED As Long
Global ROPESUSED As Long
Global MILLSTONESUSED As Long
Global PARTS_FINISHED As Long
Global PARTS_TODO As Long
Global DONE As String
Global TPOSITION As Long
Global NEWCONSTRUCTION As String
Global HEX_N As String
Global HEX_NE As String
Global HEX_SE As String
Global HEX_S As String
Global HEX_SW As String
Global HEX_NW As String
Global ROAD_N As String
Global ROAD_NE As String
Global ROAD_NW As String
Global ROAD_S As String
Global ROAD_SE As String
Global ROAD_SW As String
Global TOTAL_WORKERS As Long
Global NEW_CONSTRUCTION As String
Global SEQ_NUMBER As String
Global STOP_CONSTRUCTION As String
Global DITCH As Long
Global MOAT As Long
Global DOUBLE_MOAT As Long
Global CONCRETERS As Long
Global ENGINEERS As Long
Global WALL_WORKERS As Long
Global LOG_WALL As Long
Global STONE10 As Long
Global STONE15 As Long
Global STONE20 As Long
Global STONE25 As Long
Global STONE30 As Long
Global SINGLE_CONSTRUCTION As String
Global sCheckResult As String

' Error Handling Variables
Global errorstring As String
Global Function_Name As String
Global Function_Section As String
Global sklevel As Integer
Global RATING As Double
Global RECORD_COUNT As Long
Global Counter As Long
Global COST_CLAN As String
Global DICE As Long
Global DICE1 As Long
Global DICE2 As Long
Global GOOD As String
Global BUY_PRICE As Double
Global BUY_LIMIT As Double
Global SELL_PRICE As Double
Global SELL_LIMIT As Double
Global CURRENT_TIME As Variant
Global HOURS As Long
Global SECONDS As Long
Global RANDOM_TIME As Long
Global CLAN1 As String
Global TRIBE1 As String
Global CLAN2 As String
Global TRIBE2 As String
Global ANIMAL As String
Global Response As String
Global RECORD_DELETED As String
Global SPACE_POS As Long
Global CURRENT_COST_CLAN As String
Global RELIGION As String
Global CULT As String
Global Tribe_Checking_Provs As Long
Global Tribe_Checking_People As Long
Global Tribe_Checking_Hex As String

Global TVMWKSPACE As Workspace
Global ctl As Control
Global USE_SCREEN As String
Global Scouting As String
Global SHORTTERRAIN As String
Global SPOSITION As Long
Global SCOUTMOVEMENT As String
Global MOVEMENT_POINTS As Long
Global FLEET As String
Global SPYGLASSES As String
Global TERRAIN As String
Global NE_TERRAIN As String
Global SE_TERRAIN As String
Global S_TERRAIN As String
Global SW_TERRAIN As String
Global NW_TERRAIN As String
Global N_TERRAIN As String
Global GROUP_MOVE As String
Global ORIG_Direction As String
Global PREVIOUS_Direction As String
Global Direction As String
Global MOVEMENT_ORDERS(35) As String
Global MOVEMENT_ITERATIONS As Integer
Global MOVEMENT_LINE As String
Global MOVE_CLAN As String
Global MOVE_TRIBE As String
Global Follow_Tribe As String
Global SKILL_MOVE_TRIBE As String
Global CURRENT_MAP As String
Global NO_MOVEMENT_REASON As String
Global TRIBE_MOVEMENT1 As String
Global TRIBE_MOVEMENT2 As String
Global WIND_FIGURED_IN_MOVEMENT_COST As String
Global TM_POS As String                ' TO IDENTIFY THE CURRENT POSITION IN THE CODE
Global CSF_POS As String               ' TO IDENTIFY THE CURRENT POSITION IN THE CODE
Global Total_People As Long
Global HORSES As Long
Global HORSES_USED As Long
Global SCOUTS_USED As Long
Global WHICH_SCOUT As String
Global Elephants As Long
Global ELEPHANTS_USED As Long
Global CAMELS As Long
Global CAMELS_USED As Long
Global Number_Of_People_Mounted As Long
Global GROUP_MOUNTED As String
Global wind As String
Global WIND_DIRECTION As String
Global SHIP_TYPE As String
Global MOVEMENT_COUNT As Long
Global SCOUT_MISSION As String
Global TURN_CURRENT As String
Global SCOUT_NUMBER As Long
Global codetrack As Long
Global crlf As String
Global SURROUNDING_TERRAIN As String
Global TERRAIN_SURROUNDING_FLEET As String
Global ABLE_TO_ROW As String
Global SLOWEST_SHIP As String
Global MOVEMENT_POSSIBLE As String
Global ROWING_MOVEMENT As Long
Global SAILING_MOVEMENT As Long
Global SAILING_POSSIBLE As String
Global ROWING_POSSIBLE As String
Global SAILING_ONLY As String
Global ROWING_ONLY As String
Global MOVEMENT_ONLY As String
'Global DEATH_CHANCE As Long
Global DEATH_CHANCE As Double
Global pllevel As Long
Global NUMBER_OF_HERBS As Long
Global NUMBER_OF_HIVES As Long
Global NUMBER_OF_WAX As Long
Global NUMBER_OF_HONEY As Long
Global NUMBER_OF_SPICE As Long
Global NUMBER_OF_MINERALS As Long
Global NUMBER_OF_GOATS As Long
Global Number_Of_Cattle As Long
Global Number_Of_Horses As Long
Global NUMBER_OF_SHEEP As Long
Global Number_Of_Camels As Long
Global Number_Of_Elephants As Long
Global NUMBER_OF_DOGS As Long
Global NUMBER_OF_WAGONS As Long
Global ANIMAL_FIND As String
Global AMOUNT_OF_FINDS As Long
Global cnt1 As Long
Global cnt2 As Long
Global Find_Roll As Long
Global NUMBER_OF_ITEMS As Long
Global Item_Roll As Long
Global SURROUNDING_HEX(6) As String
Global NEW_HEX_N As String
Global NEW_HEX_NE As String
Global NEW_HEX_SE As String
Global NEW_HEX_S As String
Global NEW_HEX_SW As String
Global NEW_HEX_NW As String
Global NEW_HEX_NN As String
Global NEW_HEX_NNE As String
Global NEW_HEX_NENE As String
Global NEW_HEX_NESE As String
Global NEW_HEX_SESE As String
Global NEW_HEX_SSE As String
Global NEW_HEX_SS As String
Global NEW_HEX_SSW As String
Global NEW_HEX_SWSW As String
Global NEW_HEX_SWNW As String
Global NEW_HEX_NWNW As String
Global NEW_HEX_NNW As String
Global COAST_N As String, COAST_NE As String, COAST_SE As String
Global COAST_S As String, COAST_SW As String, COAST_NW As String
Global FORD_N As String, FORD_NE As String, FORD_SE As String
Global FORD_S As String, FORD_SW As String, FORD_NW As String
Global PASS_N As String
Global PASS_NE As String
Global PASS_SE As String
Global PASS_S As String
Global PASS_SW As String
Global PASS_NW As String
Global OCEAN_N As String
Global OCEAN_NE As String
Global OCEAN_SE As String
Global OCEAN_S As String
Global OCEAN_SW As String
Global OCEAN_NW As String
Global LAKE_N As String
Global LAKE_NE As String
Global LAKE_SE As String
Global LAKE_S As String
Global LAKE_SW As String
Global LAKE_NW As String
Global MOUNTAIN_N As String
Global MOUNTAIN_NE As String
Global MOUNTAIN_SE As String
Global MOUNTAIN_S As String
Global MOUNTAIN_SW As String
Global MOUNTAIN_NW As String
Global N_HEX As String
Global NE_HEX As String
Global SE_HEX As String
Global S_HEX As String
Global SW_HEX As String
Global NW_HEX As String
Global FIRST_N_HEX As String
Global FIRST_NE_HEX As String
Global FIRST_SE_HEX As String
Global FIRST_S_HEX As String
Global FIRST_SW_HEX As String
Global FIRST_NW_HEX As String
Global MOVEMENT_COST As Long
Global NEW_HEX As String
Global NEW_TERRAIN As String
Global PASS_AVAILABLE As String
Global START_TIME As Variant
Global MOVED As String
Global FCLAN As String
Global FTRIBE As String
Global JETTY_AVAILABLE As String
Global SAILING As String
Global SLENGTH As Long
Global SPOSTION As Long
Global END_TIME As Variant
Global ORIG_DOWN_MAP_LETTER As String
Global ORIG_ACROSS_MAP_LETTER As String
Global ORIG_ACROSS_NUMBER As Long
Global ORIG_DOWN_NUMBER As Long
Global WORK_ACROSS_NUMBER As Long
Global WORK_DOWN_NUMBER As Long
Global TRANSLATED_MAP_DOWN As String
Global TRANSLATED_MAP_ACROSS As String
Global NEW_DOWN_NUMBER As Long
Global NEW_ACROSS_NUMBER As Long
Global NEW_DOWN_MAP_LETTER As String
Global NEW_ACROSS_MAP_LETTER As String
Global MAPNUMBER As String
Global HEXNUMBER As String
Global TRIBESINHEX As String
Global TRIBESINHEX_NEW As String
Global TOTAL_TIME As Variant
Global MINERALSINHEX As String
Global NEW_ORDERS As String
Global STOP_LOOP As String
Global TRIBECHECK As Recordset
Global CLAN As String
Global TRIBE As String
Global PRIMARY_TRIBE As String
Global Scout_Movement_Allowed(8) As String
Global Scout_Direction(8) As String
Global Number_Of_Scouts(8) As Integer
Global Scouting_Movement(1 To 8, 1 To 8) As String
Global ACTIVITY_LABEL As String
Global TRIBES_WEIGHT As Double
Global TRIBES_CAPACITY As Double
Global Walking_Capacity As Long
Global TEMP_TRIBE As String
Global POLITICS_CLAN_CORRECT As String
Global MOVEFORM As Form
Global stext1 As String
Global stext2 As String
Global stext3 As String
Global stext4 As String
Global stext5 As String
Global stext6 As String
Global stext7 As String
Global stext8 As String
Global stext9 As String
Global stext10 As String
Global stext11 As String
Global stext12 As String
Global stext13 As String
Global stext14 As String
Global Truced_Clans As String


Function A100_turn_activities()
On Error GoTo ERR_A100_TURN_ACTIVITIES

TRIBE_STATUS = "A100_turn_activities"
SECTION_NAME = "ACTIVITIES"
DebugOP "A100_turn_activities"

'SET WORKSPACE

Call A150_Open_Tables("ALL")

DoCmd.Hourglass True   ' This command will turn the hourglass symbol on.

' Need to identify who is being worked on
' Open Activity_Sequence table and position at the first one.

If ACTSEQTAB.EOF Then
   Exit Function
Else
   
   CURRENT_ACTIVITY = ACTSEQTAB![ACTIVITY]
   Do ' This loop will loop through the Activity_Sequence table until all activities are
      ' completed.
      PROCESSACTIVITY.MoveFirst
       
      Do  ' This loop will loop through all Activities in the Process_Tribes_Activity
          Output_Processed = "No"
          ' table until all Tribe Activities for the Current_Activity are processed.
          If PROCESSACTIVITY![PROCESSED] = "Y" Then
             PROCESSACTIVITY.MoveNext
          ElseIf PROCESSACTIVITY![ACTIVITY] = CURRENT_ACTIVITY Then
             TRIBE_STATUS = "Processing Activities"
             
             TTRIBENUMBER = PROCESSACTIVITY![TRIBE]
             TCLANNUMBER = "0" & Mid(TTRIBENUMBER, 2, 3)
             'If TTRIBENUMBER is not a tribe then Skill_Tribe must be
             If Len(TTRIBENUMBER) > 4 Then
                Skill_Tribe = Left(TTRIBENUMBER, 4)
             Else
                Skill_Tribe = TTRIBENUMBER
             End If

             Call A150_INITIALISE
             If Not IsNull(PROCESSACTIVITY![ACTIVITY]) Then
                TActivity = PROCESSACTIVITY![ACTIVITY]
             Else
                TActivity = "NONE"
             End If
             If Not IsNull(PROCESSACTIVITY![ITEM]) Then
                TItem = PROCESSACTIVITY![ITEM]
             Else
                TItem = "NONE"
             End If
             If Not IsNull(PROCESSACTIVITY![DISTINCTION]) Then
                TDistinction = PROCESSACTIVITY![DISTINCTION]
             Else
                TDistinction = "NONE"
             End If
             If Not IsNull(PROCESSACTIVITY![JOINT]) Then
                TJoint = PROCESSACTIVITY![JOINT]
             Else
                TJoint = "N"
             End If
             If Not IsNull(PROCESSACTIVITY![PEOPLE]) Then
                TActives = PROCESSACTIVITY![PEOPLE]
             Else
                TActives = 0
             End If
             If Not IsNull(PROCESSACTIVITY![Slaves]) Then
                TSlaves = PROCESSACTIVITY![Slaves]
             Else
                TSlaves = 0
             End If
             If Not IsNull(PROCESSACTIVITY![SPECIALISTS]) Then
                TSpecialists = PROCESSACTIVITY![SPECIALISTS]
             Else
                TSpecialists = 0
             End If
             If Not IsNull(PROCESSACTIVITY![OWNING_TRIBE]) Then
                TOwning_Tribe = PROCESSACTIVITY![OWNING_TRIBE]
             Else
                TOwning_Tribe = TTRIBENUMBER
             End If
                
             TOwning_Clan = "0" & Mid(TOwning_Tribe, 2, 3)
             
             If Not IsNull(PROCESSACTIVITY![Number_of_Seeking_Groups]) Then
                Number_of_Seeking_Groups = PROCESSACTIVITY![Number_of_Seeking_Groups]
             Else
                Number_of_Seeking_Groups = 0
             End If
             If Not IsNull(PROCESSACTIVITY![Whale_Size]) Then
                TWhale_Size = PROCESSACTIVITY![Whale_Size]
             Else
                TWhale_Size = "S"
             End If
             If Not IsNull(PROCESSACTIVITY![MINING_DIRECTION]) Then
                TMining_Direction = PROCESSACTIVITY![MINING_DIRECTION]
             Else
                TMining_Direction = "NONE"
             End If
             If Not IsNull(PROCESSACTIVITY![Building]) Then
                TBuilding = PROCESSACTIVITY![Building]
             Else
                TBuilding = 0
             End If
             
             ' Add slaves to actives so that correct number of workers is employed
             TActives = TActives + TSlaves
             
             Call A150_GET_ACTIVITY_RECORD
             If No_Activity_Record_Found = "No" Then
                GoTo Finish_Processing_Activity
             End If
             Call A250_Get_Tribe_Info
             If INVALID_TRIBE = "Y" Then
                GoTo Finish_Processing_Activity
             End If
             Call A250_Get_Hex_Info
             Call A250_Get_Weather_Info
             Call A250_Get_Tribes_Skills
             Call A250_GET_SEASON_DATA
             Call A250_GET_TERRAIN_DATA
             Call A250_GET_WIND_DATA
             Call A250_PERFORM_SKILLS_CHECK
             
             ' Process a dice roll prior to starting the main process
             roll1 = DROLL(6, 1, 100, 5, DICE_TRIBE, 1, 0)
            
             Call A500_MAIN_PROCESS
             
Finish_Processing_Activity:
             Call A800_OUTPUT_PROCESSING
             Output_Processed = "Yes"
             ' need to update the processed flag.
             ' DO NOT NEED TO IF RECORD BEING PROCESSED WAS A DEFAULT TURN
             If TActivity = "DEFAULT" Then
                 Call A150_Open_Tables("ALL")
                 PROCESSACTIVITY.MoveFirst
             Else
                 PROCESSACTIVITY.Edit
                 PROCESSACTIVITY![PROCESSED] = "Y"
                 PROCESSACTIVITY.UPDATE
             End If
             
'            need to update a table at end of activity, the table would need to be setup
'            at start of turn and updated when transfers occur.
'            this would be used to update the amount of actives that have been used against
'            what was available, like implements, goods, etc.
'            TActivesAvailable = TActivesAvailable - TActives

             PROCESSACTIVITY.MoveNext
          Else
             PROCESSACTIVITY.MoveNext
          End If
INVALID_TRIBE_JUMP_POINT:
          If PROCESSACTIVITY.EOF Then
             Exit Do
          End If
      Loop
      ACTSEQTAB.MoveNext
      If ACTSEQTAB.EOF Then
         DoCmd.Hourglass False
         Exit Function
      Else
         CURRENT_ACTIVITY = ACTSEQTAB![ACTIVITY]
      End If
   Loop
End If
             
Call A900_Close_Tables

DebugOP "FINISHED - A100_turn_activities)"

ERR_A100_TURN_ACTIVITIES_CLOSE:
   Exit Function


ERR_A100_TURN_ACTIVITIES:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A100_TURN_ACTIVITIES_CLOSE
End If

End Function

Public Function Calc_New_Num_Occurs()
On Error GoTo ERR_Calc_New_Num_Occurs
TRIBE_STATUS = "Calc_New_Num_Occurs"

Dim TEST1 As Long
Dim TEST2 As Double

   TQuantity(Index1) = TQuantity(Index1) / TNUMOCCURS
   TEST1 = CLng(TEMPITEM / TQuantity(Index1))
   TEST2 = (TEMPITEM / TQuantity(Index1))
   If TEST1 > TEST2 Then
      TNUMOCCURS = CLng(TEMPITEM / TQuantity(Index1)) - 1
   Else
      TNUMOCCURS = CLng(TEMPITEM / TQuantity(Index1))
   End If
   NumItemsMade = TNUMOCCURS * TNumItems
   TQuantity(Index1) = TQuantity(Index1) * TNUMOCCURS
   If Index1 > 1 Then
      Index1 = 0
   End If

ERR_Calc_New_Num_Occurs_CLOSE:
   Exit Function

ERR_Calc_New_Num_Occurs:
  If Err = 6 Then
     Exit Function
  Else
     Call A999_ERROR_HANDLING
     Resume ERR_Calc_New_Num_Occurs_CLOSE
  End If
  
End Function

Public Function Check_Research()






End Function


Function FINAL_ACTIVITIES(Pop_Growth, Slave_Growth, Eating_People, Eating_Animals, Other)
On Error GoTo ERR_FINAL_ACTIVITIES
TRIBE_STATUS = "FINAL_ACTIVITIES"
EXECUTION_STATUS = "Start"
Dim Available_Silver As Long
Dim Required_Silver As Long
Dim Mercenaries_Unpaid As Long
Dim Hirelings_Unpaid As Long

Forms![FINAL_ACTIVITIES]![Status] = "Starting Final Activties"
Forms![FINAL_ACTIVITIES].Repaint

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

If GMTABLE![FINAL_ACTIVITIES_PROCESSED] = "Y" Then
   Msg = "Final Activities has already been processed!!!"
   MsgBox (Msg)
   Exit Function
End If

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

EXECUTION_STATUS = "Open Tables"
Call A150_Open_Tables("ALL")

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst

Current_Turn = globalinfo![CURRENT TURN]
TURN_NUMBER = "TURN" & Left(globalinfo![CURRENT TURN], 2)

MAIN_GROUP = "0330"
MAIN_INCREASE = 10

Do Until TRIBESGOODS.EOF
   If IsNull(TRIBESGOODS![ITEM_NUMBER]) Then
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
   End If
   TRIBESGOODS.MoveNext
   If TRIBESGOODS.EOF Then
      Exit Do
   End If
Loop

'need to commence process from a clan until another clan
TO_CLANNUMBER = Forms![FINAL_ACTIVITIES]![TO_CLANNUMBER]
FROM_CLANNUMBER = Forms![FINAL_ACTIVITIES]![FROM_CLANNUMBER]

TRIBESINFO.Seek "=", FROM_CLANNUMBER, FROM_CLANNUMBER

Do Until TRIBESINFO.EOF
   TRIBE_STATUS = "FINAL_ACTIVITIES"
   EXECUTION_STATUS = "Start of do"
   TCLANNUMBER = TRIBESINFO![CLAN]
   TTRIBENUMBER = TRIBESINFO![TRIBE]
   Tribes_Current_Hex = TRIBESINFO![CURRENT HEX]
   TRIBES_TERRAIN = TRIBESINFO![CURRENT TERRAIN]

Forms![FINAL_ACTIVITIES]![Status] = "Processing Tribe" & TCLANNUMBER & " " & TTRIBENUMBER
Forms![FINAL_ACTIVITIES].Repaint

TurnActOutPut = "^BFinal Activities:^B "
   
Available_Silver = 0
Required_Silver = 0
Mercenaries_Unpaid = 0
Hirelings_Unpaid = 0
   
   ' Get Tribes Skills
   EXECUTION_STATUS = "Get Skills"
   If Len(TTRIBENUMBER) > 4 Then
      Skill_Tribe = Left(TTRIBENUMBER, 4)
   Else
      Skill_Tribe = TTRIBENUMBER
   End If
      
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   If TRIBESINFO.NoMatch Then
      Msg = TTRIBENUMBER & " IS MISSING "
      MsgBox (Msg)
      MORALE = 1
   Else
      MORALE = TRIBESINFO![MORALE]
   End If
   
   ' this only gets the skills for the current tribe, not everyone in the hex
   ' this is only true for population growth i believe
   Meeting_House_Hex = Tribes_Current_Hex
   Forms![FINAL_ACTIVITIES]![Status] = "Get Skills for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   Call CHECK_FOR_BUILDING("MEETING HOUSE")
   If BUILDING_FOUND = "Y" Then
      Call A250_Get_Tribes_Skills
      TRIBESINHEX = WHO_IS_IN_HEX(TCLANNUMBER, TTRIBENUMBER, Tribes_Current_Hex, "N")
      count = 0
      Do
         String_Length = Len(TRIBESINHEX)
         BRACKET = InStr(TRIBESINHEX, ",")
         If BRACKET = 0 Then
            TRIBE = TRIBESINHEX
            TRIBESINHEX = "EMPTY"
         Else
            TRIBE = Mid(TRIBESINHEX, 1, BRACKET - 1)
         End If
         SKILLSTABLE.MoveFirst
         SKILLSTABLE.Seek "=", TRIBE, "HEALING"
         If SKILLSTABLE.NoMatch Then
            'ignore
         ElseIf HEALING_LEVEL >= SKILLSTABLE![SKILL LEVEL] Then
            ' IGNORE
         Else
            HEALING_LEVEL = SKILLSTABLE![SKILL LEVEL]
         End If
         SKILLSTABLE.MoveFirst
         SKILLSTABLE.Seek "=", TRIBE, "SANITATION"
         If SKILLSTABLE.NoMatch Then
            'ignore
         ElseIf SANITATION_LEVEL >= SKILLSTABLE![SKILL LEVEL] Then
            ' IGNORE
         Else
            SANITATION_LEVEL = SKILLSTABLE![SKILL LEVEL]
         End If
         If TRIBESINHEX = "EMPTY" Then
            Exit Do
         Else
            TRIBESINHEX = Mid(TRIBESINHEX, (BRACKET + 2), String_Length - BRACKET)
         End If
      Loop
      
   Else
      Call A250_Get_Tribes_Skills
   End If
  
   'Determine Goods_Tribe
   EXECUTION_STATUS = "Get Goods Tribe"

   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      GOODS_TRIBE = TTRIBENUMBER
   End If

   Forms![FINAL_ACTIVITIES]![Status] = "Process Other Activties for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Other Final Activities"
   
   If Other = "Y" Then
      Call Process_Other_Final_Activities
   End If

   Forms![FINAL_ACTIVITIES]![Status] = "Process Population Growth for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Population Growth"
   If Pop_Growth = "Y" Then
      Call Process_Population_Growth
   End If

   Forms![FINAL_ACTIVITIES]![Status] = "Process Slave Growth for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Slave Growth"
   If Slave_Growth = "Y" Then
      Call Process_Slave_Growth
   End If

   Forms![FINAL_ACTIVITIES]![Status] = "Process Eating for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   If TCLANNUMBER <> "0263" Then
   
   EXECUTION_STATUS = "Eating time"
   If Eating_People = "Y" Then
      If Left(GOODS_TRIBE, 1) = "B" Then
         ' BANDITS DON'T HAVE TO WORRY ABOUT EATING.
      Else
         'perform Other_final_activities
         Call Process_People_Eating
      End If
   End If
   Forms![FINAL_ACTIVITIES]![Status] = "Process Drinking for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Drinking time"
   If Left(GOODS_TRIBE, 1) = "B" Then
         ' BANDITS DON'T HAVE TO WORRY ABOUT DRINKING.
   Else
         'perform Other_final_activities
         Call Process_People_Drinking
   End If
   Forms![FINAL_ACTIVITIES]![Status] = "Process Eating Animals for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Eating Animals"
   If Eating_Animals = "Y" Then
       If Left(GOODS_TRIBE, 1) = "B" Then
         ' BANDITS DON'T HAVE TO WORRY ABOUT ANIMALS EATING.
      Else
         'perform Other_final_activities
         Call Process_Animal_Eating
      End If
   End If
   Forms![FINAL_ACTIVITIES]![Status] = "Process Mercenaries for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Paying Mercenaries"
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   
      If TRIBESINFO![MERCENARIES] > 0 Then
         ' if silver >= mercanaries * 10 all good
         Available_Silver = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "SILVER")
         Required_Silver = TRIBESINFO![MERCENARIES] * 10
         
         If Available_Silver = 0 Then
            Call Check_Turn_Output("", " There is insufficient silver to pay Mercenaries this turn.", "", 0, "NO")
            Msg = GOODS_TRIBE & " IS MISSING " & Required_Silver & " Silver for Mercenaries"
            MsgBox (Msg)
         ElseIf Required_Silver > Available_Silver Then
            Call Check_Turn_Output("", " There is insufficient silver to pay Mercenaries this turn.", "", 0, "NO")
            Msg = GOODS_TRIBE & "IS MISSING " & (Required_Silver - Available_Silver) & "Silver to pay Mercenaries"
            MsgBox (Msg)
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - Available_Silver
            TRIBESGOODS.UPDATE
            ' Reduce Mercenaries by 50% of those unpaid
            Mercenaries_Unpaid = (Required_Silver - Available_Silver) / 10
            TRIBESINFO.MoveFirst
            TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
            TRIBESINFO.Edit
            TRIBESINFO![MERCENARIES] = TRIBESINFO![MERCENARIES] - (Mercenaries_Unpaid / 5)
            TRIBESINFO.UPDATE
         Else
            ' Availble silver matches or exceed silver required
            If TTRIBENUMBER = "0330" Then
               Msg = "Silver before Mercenaries are deducted is " & Available_Silver & ", "
               Call Check_Turn_Output("", Msg, "", 0, "NO")
            End If
            Msg = "Mercenaries cost " & Required_Silver & " silver. "
            Call Check_Turn_Output("", Msg, "", 0, "NO")
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - Required_Silver
            TRIBESGOODS.UPDATE
         End If
      End If
      
   
   'Pay Hirelings
   Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TTRIBENUMBER, "ACTIVITIES", 2, TurnActOutPut, "No") '?
   
   EXECUTION_STATUS = "Paying Hirelings"
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   
      If TRIBESINFO![HIRELINGS] > 0 Then
         ' if silver >= Hirelings * 10 all good
         Available_Silver = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "SILVER")
         Required_Silver = TRIBESINFO![HIRELINGS] * 10
         
         If Available_Silver = 0 Then
            Call Check_Turn_Output("", " There is insufficient silver to pay Hirelings this turn.", "", 0, "NO")
            Msg = GOODS_TRIBE & " IS MISSING " & Required_Silver & " Silver for Hirelings"
            MsgBox (Msg)
         ElseIf Required_Silver > Available_Silver Then
            Call Check_Turn_Output("", " There is insufficient silver to pay Hirelings this turn.", "", 0, "NO")
            Msg = GOODS_TRIBE & "IS MISSING " & (Required_Silver - Available_Silver) & "Silver to pay Hirelings"
            MsgBox (Msg)
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - Available_Silver
            TRIBESGOODS.UPDATE
            ' Reduce Hirelings by 50% of those unpaid
            Hirelings_Unpaid = (Required_Silver - Available_Silver) / 10
            TRIBESINFO.MoveFirst
            TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
            TRIBESINFO.Edit
            TRIBESINFO![HIRELINGS] = TRIBESINFO![HIRELINGS] - (Hirelings_Unpaid / 5)
            TRIBESINFO.UPDATE
         Else
            ' Availble silver matches or exceed silver required
            If TTRIBENUMBER = "0330" Then
               Msg = "Silver before Hirelings are deducted is " & Available_Silver & ", "
               Call Check_Turn_Output("", Msg, "", 0, "NO")
            End If
            Msg = "Hirelings cost " & Required_Silver & " silver. "
            Call Check_Turn_Output("", Msg, "", 0, "NO")
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - Required_Silver
            TRIBESGOODS.UPDATE
         End If
      End If
      
   End If
   
   'Process all final activities prior to politics
   Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TTRIBENUMBER, "ACTIVITIES", 2, TurnActOutPut, "No")
   
   Forms![FINAL_ACTIVITIES]![Status] = "Process Politics for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   EXECUTION_STATUS = "Government stuff"
   Call Initialise_Politics_Variables

   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   Government_Level = TRIBESINFO![GOVT LEVEL]
   ' PERFORM PACIFICATION ATTEMPT
   ' Perform Politics related activities
   ' get Government level
   ' get pl levels for each hex
   ' perform activities for each hex
   ' update activities for each hex
   TurnActOutPut = "^BPolitical Tithes:^B "
   String_Length = Len(TTRIBENUMBER)
   If String_Length > 4 Then
      ' DO NOTHING
   ElseIf POLITICS_LEVEL >= 10 Then
      ' GET PACIFICATION LEVELS FOR HEXES BEING CONTROLLED
      Starting_Hex_GL0 = Tribes_Current_Hex
      CURRENT_HEX = Tribes_Current_Hex
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", CURRENT_HEX
      CURRENT_TERRAIN = hexmaptable![TERRAIN]
      PACIFICATION_TABLE.Seek "=", TCLANNUMBER, TTRIBENUMBER
      If PACIFICATION_TABLE.NoMatch Then
          PACIFICATION_TABLE.AddNew
          PACIFICATION_TABLE![CLAN] = TCLANNUMBER
          PACIFICATION_TABLE![TRIBE] = TTRIBENUMBER
          PACIFICATION_TABLE.UPDATE
          PACIFICATION_TABLE.Seek "=", TCLANNUMBER, TTRIBENUMBER
      End If
      
      ' 1 % OF WARRIORS PER PACIFICATION ATTEMPT
            
            
            
      If Government_Level >= 0 Then
         OutLine = " Primary Hex - tithes : "
         If PACIFICATION_TABLE![primary_hex] >= 20 Then
             Call PERFORM_PACIFICATION(CURRENT_HEX, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
         End If
         Call Perform_Politics_Activities("GL0")
      End If
      
      MOVEMENT(1) = "NONE"
      MOVEMENT(2) = "NONE"
      MOVEMENT(3) = "NONE"
      MOVEMENT(4) = "NONE"
      MOVEMENT(5) = "NONE"
      MOVEMENT(6) = "NONE"
      
      
      ' OK - new process
      Iteration_Count(1) = 1
      If Government_Level >= 1 Then
         Do Until Iteration_Count(1) > 6
            MOVEMENT(1) = Ring1(Iteration_Count(1))
         
            Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, MOVEMENT(1), MOVEMENT(2), MOVEMENT(3), MOVEMENT(4), MOVEMENT(5), MOVEMENT(6), "NONE", "NONE")
         
            OutLine = " Hex to " & MOVEMENT(1) & " - tithes : "
            If PACIFICATION_TABLE![GL1_1] >= 20 Then  ' this is incorrect, should be incrementing field
               Call PERFORM_PACIFICATION(CURRENT_HEX, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
            End If
            Call Perform_Politics_Activities("GL1")
            Iteration_Count(1) = Iteration_Count(1) + 1
         Loop
      End If
      
      Iteration_Count(1) = 1
      Iteration_Count(2) = 1
      MOVEMENT(1) = "NONE"
      MOVEMENT(2) = "NONE"
      
      If Government_Level >= 2 Then
         Do Until Iteration_Count(1) > 12
            ' for each iteration, read the two movements from each Ring2 variable
            Do Until Iteration_Count(2) > 2
               BRACKET = 0
               BRACKET = InStr(Ring2(Iteration_Count(1)), ",")
               If BRACKET > 0 Then
                  MOVEMENT(Iteration_Count(2)) = Left(Ring2(Iteration_Count(1)), (BRACKET - 1))
               Else
                  MOVEMENT(Iteration_Count(2)) = Ring2(Iteration_Count(1))
               End If
               Ring2(Iteration_Count(1)) = Right(Ring2(Iteration_Count(1)), (Len(Ring2(Iteration_Count(1))) - BRACKET))
               Iteration_Count(2) = Iteration_Count(2) + 1
            Loop
            ' now call pacification code
            Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, MOVEMENT(1), MOVEMENT(2), MOVEMENT(3), MOVEMENT(4), MOVEMENT(5), MOVEMENT(6), "NONE", "NONE")
            OutLine = " Hex to " & MOVEMENT(1) & "/" & MOVEMENT(2) & " - tithes : "
            If PACIFICATION_TABLE![GL1_1] >= 20 Then  ' this is incorrect, same issue as GL1
               Call PERFORM_PACIFICATION(CURRENT_HEX, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
            End If
            Call Perform_Politics_Activities("GL2")
            Iteration_Count(2) = 1
            Iteration_Count(1) = Iteration_Count(1) + 1
         Loop
      End If
     
      Iteration_Count(1) = 1
      Iteration_Count(2) = 1
      MOVEMENT(1) = "NONE"
      MOVEMENT(2) = "NONE"
      MOVEMENT(3) = "NONE"
      
      If Government_Level >= 3 Then
         Do Until Iteration_Count(1) > 18
            ' for each iteration, read the two movements from each Ring2 variable
            Do Until Iteration_Count(2) > 3
               BRACKET = 0
               BRACKET = InStr(Ring3(Iteration_Count(1)), ",")
               If BRACKET > 0 Then
                  MOVEMENT(Iteration_Count(2)) = Left(Ring3(Iteration_Count(1)), (BRACKET - 1))
               Else
                  MOVEMENT(Iteration_Count(2)) = Ring3(Iteration_Count(1))
               End If
               Ring3(Iteration_Count(1)) = Right(Ring3(Iteration_Count(1)), (Len(Ring3(Iteration_Count(1))) - BRACKET))
               Iteration_Count(2) = Iteration_Count(2) + 1
            Loop
            ' now call pacification code
            Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, MOVEMENT(1), MOVEMENT(2), MOVEMENT(3), MOVEMENT(4), MOVEMENT(5), MOVEMENT(6), "NONE", "NONE")
            OutLine = " Hex to " & MOVEMENT(1) & "/" & MOVEMENT(2) & "/" & MOVEMENT(3) & " - tithes : "
            If PACIFICATION_TABLE![GL1_1] >= 20 Then
               Call PERFORM_PACIFICATION(CURRENT_HEX, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
            End If
            Call Perform_Politics_Activities("GL3")
            Iteration_Count(2) = 1
            Iteration_Count(1) = Iteration_Count(1) + 1
         Loop
      End If
    
      Iteration_Count(1) = 1
      Iteration_Count(2) = 1
      MOVEMENT(1) = "NONE"
      MOVEMENT(2) = "NONE"
      MOVEMENT(3) = "NONE"
      MOVEMENT(4) = "NONE"
      
      If Government_Level >= 4 Then
         Do Until Iteration_Count(1) > 24
            ' for each iteration, read the two movements from each Ring2 variable
            Do Until Iteration_Count(2) > 4
               BRACKET = 0
               BRACKET = InStr(Ring4(Iteration_Count(1)), ",")
               If BRACKET > 0 Then
                  MOVEMENT(Iteration_Count(2)) = Left(Ring4(Iteration_Count(1)), (BRACKET - 1))
               Else
                  MOVEMENT(Iteration_Count(2)) = Ring4(Iteration_Count(1))
               End If
               Ring4(Iteration_Count(1)) = Right(Ring4(Iteration_Count(1)), (Len(Ring4(Iteration_Count(1))) - BRACKET))
               Iteration_Count(2) = Iteration_Count(2) + 1
            Loop
            ' now call pacification code
            Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, MOVEMENT(1), MOVEMENT(2), MOVEMENT(3), MOVEMENT(4), MOVEMENT(5), MOVEMENT(6), "NONE", "NONE")
            OutLine = " Hex to " & MOVEMENT(1) & "/" & MOVEMENT(2) & "/" & MOVEMENT(3) & "/" & MOVEMENT(4) & " - tithes : "
            If PACIFICATION_TABLE![GL1_1] >= 20 Then
               Call PERFORM_PACIFICATION(CURRENT_HEX, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
            End If
            Call Perform_Politics_Activities("GL4")
            Iteration_Count(2) = 1
            Iteration_Count(1) = Iteration_Count(1) + 1
         Loop
      End If
    
'      Iteration_Count(1) = 1
'      Iteration_Count(2) = 1
'      Movement(1) = "NONE"
'      Movement(2) = "NONE"
'      Movement(3) = "NONE"
'      Movement(4) = "NONE"
'      Movement(5) = "NONE"
'
'      If Government_Level >= 5 Then
'         Do Until Iteration_Count(1) > 30
'            ' for each iteration, read the two movements from each Ring2 variable
'            Do Until Iteration_Count(2) > 5
'               BRACKET = 0
'               BRACKET = InStr(Ring5(Iteration_Count(1)), ",")
'               If BRACKET > 0 Then
'                  Movement(Iteration_Count(2)) = Left(Ring5(Iteration_Count(1)), (BRACKET - 1))
'               Else
'                  Movement(Iteration_Count(2)) = Ring5(Iteration_Count(1))
'               End If
'               Ring5(Iteration_Count(1)) = Right(Ring5(Iteration_Count(1)), (Len(Ring5(Iteration_Count(1))) - BRACKET))
'               Iteration_Count(2) = Iteration_Count(2) + 1
'            Loop
'            ' now call pacification code
'            Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, Movement(1), Movement(2), Movement(3), Movement(4), Movement(5), Movement(6), "NONE", "NONE")
'            OutLine = " Hex to " & Movement(1) & "/" & Movement(2) & "/" & Movement(3) & "/" & Movement(4)
'            OutLine = OutLine & "/" & Movement(5) & " - tithes : "
'            If PACIFICATION_TABLE![GL1_1] >= 20 Then
'               Call PERFORM_PACIFICATION(Current_Hex, CURRENT_TERRAIN, TCLANNUMBER, TTRIBENUMBER)
'            End If
'            Call Perform_Politics_Activities("GL5")
'            Iteration_Count(2) = 1
'            Iteration_Count(1) = Iteration_Count(1) + 1
'         Loop
'      End If
       
       If Len(TurnActOutPut) > 0 Then
          Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TTRIBENUMBER, "ACTIVITIES", 3, TurnActOutPut, "No")
       End If
   End If

   EXECUTION_STATUS = "Annual Costs"
   ' Only process in turn 01
   Forms![FINAL_ACTIVITIES]![Status] = "Process Annual Costs for " & TTRIBENUMBER
   Forms![FINAL_ACTIVITIES].Repaint

   Forms![FINAL_ACTIVITIES]![Status] = "Move to next unit"
   Forms![FINAL_ACTIVITIES].Repaint


   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.MoveNext
   If TRIBESINFO![CLAN] <> TCLANNUMBER Then
      ' need to process Research costs for last clan
      Call Process_Clan_Research_Costs
      ' Nullify all unstable items
    TRIBESGOODS.MoveFirst
    Do Until TRIBESGOODS.EOF
        If TRIBESGOODS![CLAN] = TCLANNUMBER Then
            If TRIBESGOODS![ITEM] = "FISH" Or _
                TRIBESGOODS![ITEM] = "MILK" Or _
                TRIBESGOODS![ITEM] = "BREAD" Then
                TRIBESGOODS.Edit
                TRIBESGOODS![ITEM_NUMBER] = 0
                TRIBESGOODS.UPDATE
            End If
        End If
        TRIBESGOODS.MoveNext
    Loop

   End If
   
   If TRIBESINFO.EOF Then
      Exit Do
   End If
   If TRIBESINFO![CLAN] > TO_CLANNUMBER Then
      Exit Do
   End If
Loop

Forms![FINAL_ACTIVITIES]![Status] = "Clean up of TRIBESGOODS"
Forms![FINAL_ACTIVITIES].Repaint


TRIBESGOODS.MoveFirst

Do Until TRIBESGOODS.EOF
   If TRIBESGOODS![ITEM_NUMBER] <= 0 Then
      TRIBESGOODS.Delete
   End If
   TRIBESGOODS.MoveNext
Loop

    
Forms![FINAL_ACTIVITIES]![Status] = "Clean up Specialists"
Forms![FINAL_ACTIVITIES].Repaint

' Clean Up Specialists in training
QUERY_STRING = "DELETE * FROM TRIBES_SPECIALISTS"
QUERY_STRING = QUERY_STRING & " WHERE (TRIBES_SPECIALISTS.Number_Of_Turns_Training>3);"
Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute


EXECUTION_STATUS = "Village Trading"
Forms![FINAL_ACTIVITIES]![Status] = "Village Trading"
Forms![FINAL_ACTIVITIES].Repaint


Call Village_Trading

EXECUTION_STATUS = "Capacities & Weights"
Forms![FINAL_ACTIVITIES]![Status] = "Capacities"
Forms![FINAL_ACTIVITIES].Repaint

Call POPULATE_CAPACITIES
Forms![FINAL_ACTIVITIES]![Status] = "Weights"
Forms![FINAL_ACTIVITIES].Repaint

Call POPULATE_WEIGHTS

Call A900_Close_Tables

' Update Tribe_Checking

Forms![FINAL_ACTIVITIES]![Status] = "Update Tribe Checking"
Forms![FINAL_ACTIVITIES].Repaint

Call Tribe_Checking("Update_All", "", "", "")

Forms![FINAL_ACTIVITIES]![Status] = "Finished"
Forms![FINAL_ACTIVITIES].Repaint


ERR_FINAL_ACTIVITIES_CLOSE:
   DoCmd.Hourglass False
   Exit Function

ERR_FINAL_ACTIVITIES:
TRIBE_STATUS = TRIBE_STATUS & " " & EXECUTION_STATUS
  Call A999_ERROR_HANDLING
  Resume ERR_FINAL_ACTIVITIES_CLOSE

End Function

Public Function A150_GET_ACTIVITY_RECORD()
On Error GoTo ERR_A150_GET_ACTIVITY_RECORD

TRIBE_STATUS = "A150_GET_ACTIVITY_RECORD"


    
No_Activity_Record_Found = "No"

Index1 = 1
Do Until Index1 > 40
   TGoods(Index1) = "EMPTY"
   Index1 = Index1 + 1
Loop

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", TActivity, TItem, TDistinction
If ActivitiesTable.NoMatch Then
   TurnActOutPut = TurnActOutPut & TActivity & " producing " & TItem & " using " & TDistinction & " was invalid and did not process, "
   GoTo ERR_A150_GET_ACTIVITY_RECORD_CLOSE
Else
   No_Activity_Record_Found = "Yes"
End If

TResearch = ActivitiesTable![research]
TPeople = ActivitiesTable![PEOPLE]
If Not IsNull(ActivitiesTable![SHORTNAME]) Then
   TShort = ActivitiesTable![SHORTNAME]
Else
   TShort = "EMPTY"
End If

' Set variables to do calculators
TSkilllvl(1) = ActivitiesTable![SKILL LEVEL]
If Not IsNull(ActivitiesTable![SECOND SKILL LEVEL]) Then
   TSkilllvl(2) = ActivitiesTable![SECOND SKILL LEVEL]
End If
If Not IsNull(ActivitiesTable![THIRD SKILL LEVEL]) Then
   TSkilllvl(3) = ActivitiesTable![THIRD SKILL LEVEL]
End If
If Not IsNull(ActivitiesTable![FORTH SKILL LEVEL]) Then
   TSkilllvl(4) = ActivitiesTable![FORTH SKILL LEVEL]
End If
If IsNull(ActivitiesTable![SECOND SKILL]) Then
   TSkill(2) = "FORGET"
Else
   TSkill(2) = ActivitiesTable![SECOND SKILL]
End If
If IsNull(ActivitiesTable![THIRD SKILL]) Then
   TSkill(3) = "FORGET"
Else
   TSkill(3) = ActivitiesTable![THIRD SKILL]
End If
If IsNull(ActivitiesTable![FORTH SKILL]) Then
   TSkill(4) = "FORGET"
Else
   TSkill(4) = ActivitiesTable![FORTH SKILL]
End If
If IsNull(ActivitiesTable![Item_Produced]) Then
   TItem_Produced = "FORGET"
Else
   TItem_Produced = ActivitiesTable![Item_Produced]
End If
      
TNumItems = ActivitiesTable![NUMBER OF ITEMS]

ItemsTable.index = "SECONDARYKEY"
ItemsTable.MoveFirst
ItemsTable.Seek "=", TActivity, TItem, TDistinction
    
Index1 = 1
If Not ItemsTable.NoMatch Then
   Do While (ItemsTable![ITEM]) = TItem And (ItemsTable![TYPE] = TDistinction)
      TGoods(Index1) = ItemsTable![GOOD]
      TQuantity(Index1) = ItemsTable![NUMBER]
      SQuantity(Index1) = ItemsTable![NUMBER]
      Index1 = Index1 + 1
      ItemsTable.MoveNext
      If ItemsTable.EOF Then
         Exit Do
      End If
   Loop
End If

ERR_A150_GET_ACTIVITY_RECORD_CLOSE:
   Exit Function

ERR_A150_GET_ACTIVITY_RECORD:
If (Err = 3021) Then  ' 3021 = No Current Record
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A150_GET_ACTIVITY_RECORD_CLOSE
End If
   
End Function

Public Function GET_FARMING_TURN1()
On Error GoTo ERR_GET_FARMING_TURN1
TRIBE_STATUS = "GET_FARMING_TURN1"

If Left(FARMING_TURN1, 2) = "01" Then
   FARMING_TURN1 = "02" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "02" Then
   FARMING_TURN1 = "03" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "03" Then
   FARMING_TURN1 = "04" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "04" Then
   FARMING_TURN1 = "05" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "05" Then
   FARMING_TURN1 = "06" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "06" Then
   FARMING_TURN1 = "07" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "07" Then
   FARMING_TURN1 = "08" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "08" Then
   FARMING_TURN1 = "09" & Right(Current_Turn, 4)
ElseIf Left(FARMING_TURN1, 2) = "09" Then
   FARMING_TURN1 = "10" & Right(Current_Turn, 4)
End If

ERR_GET_FARMING_TURN1_CLOSE:
   Exit Function

ERR_GET_FARMING_TURN1:
   Call A999_ERROR_HANDLING
   Resume ERR_GET_FARMING_TURN1_CLOSE

End Function

Public Function GET_FARMING_TURN2()
On Error GoTo ERR_GET_FARMING_TURN2
TRIBE_STATUS = "GET_FARMING_TURN2"

If Left(FARMING_TURN1, 2) = "01" Then
   FARMING_TURN2 = "02" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "02" Then
   FARMING_TURN2 = "03" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "03" Then
   FARMING_TURN2 = "04" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "04" Then
   FARMING_TURN2 = "05" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "05" Then
   FARMING_TURN2 = "06" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "06" Then
   FARMING_TURN2 = "07" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "07" Then
   FARMING_TURN2 = "08" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "08" Then
   FARMING_TURN2 = "09" & Right(FARMING_TURN1, 4)
ElseIf Left(FARMING_TURN1, 2) = "09" Then
   FARMING_TURN2 = "10" & Right(FARMING_TURN1, 4)
End If

ERR_GET_FARMING_TURN2_CLOSE:
   Exit Function

ERR_GET_FARMING_TURN2:
   Call A999_ERROR_HANDLING
   Resume ERR_GET_FARMING_TURN2_CLOSE

End Function

Public Function GET_FARMING_TURN3()
On Error GoTo ERR_GET_FARMING_TURN3
TRIBE_STATUS = "GET_FARMING_TURN3"

If Left(FARMING_TURN2, 2) = "01" Then
   FARMING_TURN3 = "02" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "02" Then
   FARMING_TURN3 = "03" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "03" Then
   FARMING_TURN3 = "04" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "04" Then
   FARMING_TURN3 = "05" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "05" Then
   FARMING_TURN3 = "06" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "06" Then
   FARMING_TURN3 = "07" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "07" Then
   FARMING_TURN3 = "08" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "08" Then
   FARMING_TURN3 = "09" & Right(FARMING_TURN2, 4)
ElseIf Left(FARMING_TURN2, 2) = "09" Then
   FARMING_TURN3 = "10" & Right(FARMING_TURN2, 4)
End If

ERR_GET_FARMING_TURN3_CLOSE:
   Exit Function

ERR_GET_FARMING_TURN3:
   Call A999_ERROR_HANDLING
   Resume ERR_GET_FARMING_TURN3_CLOSE

End Function

Public Function HERDING_LIMIT()
On Error GoTo ERR_HERDING_LIMIT
TRIBE_STATUS = "HERDING_LIMIT"

Dim H20_PER_HERDER As Long
Dim H10_PER_HERDER As Long
Dim H5_PER_HERDER As Long
Dim CROOKSAVAILABLE As Long
Dim FencesAvailable As Long
Dim StablesAvailable As Long
Dim Herding_Dogs_Available As Long
Dim Horse_Herder As String
Dim Mounted_Herder As String
Dim ANIMAL As String
Dim Actives_Available As Long
Dim Maximum_Herders As Long
Dim Maximum_Animals As Long
Dim Maximum_H10_Animals As Long
Dim Maximum_H20_Animals As Long
Dim Effective_Herders As Long
Dim Herd_Limit_Output As String

Dim MaxCrooks As Long  ' maximum herders required * 0.7
Dim MaxFences As Long
Dim MaxDogs As Long
Dim MaxHorses As Long
Dim MaxStables As Long
Dim Actual_Herders_Required As Long
Dim MaxHerdersRequired As Long  ' total of all animals / herding requirements
Dim HerderEffectiveness As Double  ' maximum herders required * 0.7
Dim SpecialistModifier As Double  ' Specialists / Total assigned herders
Dim FenceModifier As Double
Dim StableModifier As Double
Dim CrookRatio As Double
Dim CrookModifier As Double
Dim DogModifier As Double
Dim HorseModifier As Double
Dim Total_Horses As Long


H20_PER_HERDER = 0
H10_PER_HERDER = 0
H5_PER_HERDER = 0
HERDERSREQ = 0
FencesAvailable = 0
StablesAvailable = 0
Horse_Herder = 0
Mounted_Herder = 0
CROOKSAVAILABLE = 0

MaxCrooks = 0
MaxDogs = 0
MaxFences = 0
MaxStables = 0
Actual_Herders_Required = 0
MaxHerdersRequired = 0
HerderEffectiveness = 0
CrookRatio = 0
CrookModifier = 0
DogModifier = 0
FenceModifier = 0
StableModifier = 0
Maximum_Animals = 0


If Len(TurnActOutPut) > 20 Then
   Herd_Limit_Output = ", "
Else
   Herd_Limit_Output = " "
End If
  
' number of herders available
Actives_Available = TActives

' read valid animals and populate
' do not include herding dogs

TRIBE_STATUS = "HERDING_LIMIT Get Animals"
VALIDANIMALS.MoveFirst
ANIMAL = VALIDANIMALS![ANIMAL]
Do
  ' HERDERSREQ = HERDERSREQ + GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, Goods_Tribe, ANIMAL) / VALIDANIMALS![HERDERS_REQUIRED]

  MaxHerdersRequired = MaxHerdersRequired + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL) / VALIDANIMALS![Herders_Required])
  Maximum_Animals = Maximum_Animals + GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
  
  If VALIDANIMALS![Herders_Required] = 5 Then
     H5_PER_HERDER = H5_PER_HERDER + GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
  ElseIf VALIDANIMALS![Herders_Required] = 10 Then
     If ANIMAL = "HERDING DOG" Then
        'ignore
     ElseIf ANIMAL = "HORSE" Then
        Total_Horses = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
        H10_PER_HERDER = H10_PER_HERDER + Total_Horses
     Else
        H10_PER_HERDER = H10_PER_HERDER + GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
     End If
  ElseIf VALIDANIMALS![Herders_Required] = 20 Then
     H20_PER_HERDER = H20_PER_HERDER + GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
  End If
  VALIDANIMALS.MoveNext
  If VALIDANIMALS.EOF Then
     Exit Do
  End If
  ANIMAL = VALIDANIMALS![ANIMAL]
Loop
      
TRIBE_STATUS = "HERDING_LIMIT Get Construction"

' Check for all research & research related items & construction
' Fences or wire_fences

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, TCLANNUMBER, "FENCES"

If Not HEXMAPCONST.NoMatch Then
   If HEXMAPCONST![1] > 0 Then
      FencesAvailable = HEXMAPCONST![1]
   End If
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, TCLANNUMBER, "WIRE FENCES"

If Not HEXMAPCONST.NoMatch Then
   If HEXMAPCONST![1] > 0 Then
      FencesAvailable = FencesAvailable + HEXMAPCONST![1]
   End If
End If

If FencesAvailable > 0 Then
    Herd_Limit_Output = Herd_Limit_Output & FencesAvailable & " Fences are available, "
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, TCLANNUMBER, "STABLES"

If Not HEXMAPCONST.NoMatch Then
   If HEXMAPCONST![1] > 0 Then
      StablesAvailable = HEXMAPCONST![1]
   End If
End If

If StablesAvailable > 0 Then
    Herd_Limit_Output = Herd_Limit_Output & StablesAvailable & " Stables are available, "
End If

TRIBE_STATUS = "HERDING_LIMIT Get Research"
' require the skill tribenumber not the group number eg 030 not 030e1.

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "SMART HERDING")
        
If RESEARCH_FOUND = "Y" Then
    CROOKSAVAILABLE = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "CROOKS")
Else
    CROOKSAVAILABLE = 0
End If

If CROOKSAVAILABLE > 0 Then
    Herd_Limit_Output = Herd_Limit_Output & CROOKSAVAILABLE & " Crooks are available, "
End If

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "HORSE HERDERS")
        
If RESEARCH_FOUND = "Y" Then
   HorseModifier = Total_Horses / Maximum_Animals
   If HorseModifier > 1 Then
      HorseModifier = 1
   End If
End If

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "MOUNTED HERDERS")
        
If RESEARCH_FOUND = "Y" Then
    Mounted_Herder = "Y"
Else
    Mounted_Herder = "N"
End If

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "HERDING DOGS")
        
If RESEARCH_FOUND = "Y" Then
   Herding_Dogs_Available = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "HERDING DOG")
Else
   Herding_Dogs_Available = 0
End If

If Herding_Dogs_Available > 0 Then
    Herd_Limit_Output = Herd_Limit_Output & Herding_Dogs_Available & " Herding Dogs are available, "
End If

TRIBE_STATUS = "HERDING_LIMIT Get Specialists"
TSpecialists = 0

Call Get_Specialists_Info(TCLANNUMBER, Skill_Tribe, "HERDER")
   
SPECIALIST_FOUND = "Y"

If TSpecialists > NO_SPECIALISTS_FOUND Then
   TSpecialists = NO_SPECIALISTS_FOUND
End If


'Herders Required = MaxHerdersRequired
If TSpecialists > 0 Then
   SpecialistModifier = TSpecialists / (TActives + TSpecialists)
Else
   SpecialistModifier = 0
End If

TRIBE_STATUS = "HERDING_LIMIT Fencing Modifier"
'#Fences = AvailableFences
MaxFences = (H20_PER_HERDER / 100) + (H10_PER_HERDER / 50)
If FencesAvailable > 0 Then
   FenceModifier = FencesAvailable / MaxFences
Else
   FenceModifier = 0
End If
If FenceModifier > 1 Then
   FenceModifier = 1
End If
   
TRIBE_STATUS = "HERDING_LIMIT Stables Modifier"
MaxStables = (H20_PER_HERDER / 60) + (H10_PER_HERDER / 30)
If StableModifier > 0 Then
   StableModifier = StablesAvailable / MaxStables
Else
   StableModifier = 0
End If
If StableModifier > 1 Then
   StableModifier = 1
End If
   
TRIBE_STATUS = "HERDING_LIMIT Crooks Modifier"
MaxCrooks = MaxHerdersRequired * 0.7
If CROOKSAVAILABLE > 0 Then
   CrookRatio = CROOKSAVAILABLE / MaxCrooks
Else
   CrookRatio = 0
End If
CrookModifier = 1 - 0.3 * CrookRatio '(Also known as Smart modifier)

If CrookModifier = 1 Then
   CrookModifier = 0
End If

TRIBE_STATUS = "HERDING_LIMIT Dogs Modifier"
MaxDogs = MaxHerdersRequired / 2
If Herding_Dogs_Available > 0 Then
   DogModifier = Herding_Dogs_Available / MaxDogs
Else
   DogModifier = 0
End If
If DogModifier > 1 Then
   DogModifier = 1
End If

TRIBE_STATUS = "HERDING_LIMIT Herder Effectiveness"
HerderEffectiveness = 1 + SpecialistModifier + FenceModifier + StableModifier + DogModifier + HorseModifier
HerderEffectiveness = HerderEffectiveness + CrookModifier
'HerdersEffectiveness = (1 + 0 + 0 + 0) * 1# = 1

Actual_Herders_Required = MaxHerdersRequired / HerderEffectiveness

HERDERSREQ = Actual_Herders_Required
    
'Herd_Limit_Output = Herd_Limit_Output & MaxHerdersRequired & " Max Herders required prior to modifiers, "
Herd_Limit_Output = Herd_Limit_Output & Actual_Herders_Required & " actual herders required, "
If TCLANNUMBER = "0330" Then
   Herd_Limit_Output = Herd_Limit_Output & HerderEffectiveness & " herder effectiveness, "
End If
Call Check_Turn_Output(Herd_Limit_Output, "", "", 0, "NO")
     
ERR_HERDING_LIMIT_CLOSE:
   Exit Function

ERR_HERDING_LIMIT:
   Call A999_ERROR_HANDLING
   Resume ERR_HERDING_LIMIT_CLOSE

End Function

Public Function PERFORM_APIARISM()
On Error GoTo ERR_PERFORM_APIARISM
TRIBE_STATUS = "PERFORM_APIARISM"

Dim TNUM_HIVES As Long
Dim SPECIALIST As String

' NEED TO INSERT A CHECK FOR 1 ACTIVE PER 5 HIVES

   TNewHoney = 0
   TNewWax = 0
   TNewRoyalJelly = 0
   TNewPropolis = 0
   TNUM_HIVES = 0

' NEED TO CHECK FOR BUILDING

   Call CHECK_FOR_BUILDING("APIARY")

   If BUILDING_FOUND = "N" Then
      If Right(TurnActOutPut, 3) = "^B " Then
         TurnActOutPut = TurnActOutPut & "No apiary found"
      Else
         TurnActOutPut = TurnActOutPut & ", No apiary found"
      End If
      Exit Function
   End If
   
   TNUM_HIVES = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "HIVE")
   
   If (TNUM_HIVES / 20) > APIARYS_FOUND Then
      TNUM_HIVES = APIARYS_FOUND * 20
   End If
   
   SPECIALIST = "BEEKEEPER"
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, SPECIALIST)
   
   If NO_SPECIALISTS_FOUND = 0 Then
      SPECIALIST = "APIARIST"
      Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, SPECIALIST)
   End If
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
 
 
   ' check availability of specialist.
   If TSpecialists > 0 Then
      Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, SPECIALIST, "SPECIALISTS_USED", TSpecialists)
   End If
   
   TActives = TActives + (TSpecialists * 2)
   
   If (TActives / 5) > TNUM_HIVES Then
      TNUM_HIVES = TActives * 5
   End If
   
   TNewHoney = (TNUM_HIVES * ((10 + APIARISM_LEVEL) / 10))
   TNewHoney = CLng((TNewHoney * SEASON_HONEY) * HONEY_WEATHER)

   TNewWax = (TNUM_HIVES * ((10 + APIARISM_LEVEL) / 10))
   TNewWax = CLng((TNewWax * SEASON_WAX) * WAX_WEATHER)
       
   COMPRESTAB.MoveFirst
   COMPRESTAB.Seek "=", TTRIBENUMBER, "ROYAL JELLY"
   
   If Not COMPRESTAB.NoMatch Then
      TNewRoyalJelly = CLng(TNewWax * 0.5)
   End If

   COMPRESTAB.Seek "=", TTRIBENUMBER, "PROPOLIS"
   
   If Not COMPRESTAB.NoMatch Then
      TNewPropolis = CLng(TNewWax * 0.5)
   End If
   
   If Right(TurnActOutPut, 3) = "^B " Then
      TurnActOutPut = TurnActOutPut & "Apiary ("
   Else
      TurnActOutPut = TurnActOutPut & ", Apiary ("
   End If
   
   If TNewWax > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WAX", "ADD", TNewWax)
      Call Check_Turn_Output("", "Wax", ",", TNewWax, "NO")
   End If
   If TNewHoney > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HONEY", "ADD", TNewHoney)
      Call Check_Turn_Output(" Honey ", ", ", "", TNewHoney, "NO")
   End If
   If TNewRoyalJelly > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "ROYAL JELLY", "ADD", TNewRoyalJelly)
      Call Check_Turn_Output(" Royal Jelly ", ", ", "", TNewRoyalJelly, "NO")
   End If
   If TNewPropolis > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROPOLIS", "ADD", TNewPropolis)
      Call Check_Turn_Output(" Propolis ", ", ", "", TNewPropolis, "NO")
   End If
   DoCmd.Hourglass True
   
   TurnActOutPut = TurnActOutPut & ")"
   
ERR_PERFORM_APIARISM_CLOSE:
   Exit Function

ERR_PERFORM_APIARISM:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_APIARISM_CLOSE

End Function

Public Function PERFORM_APOTHECARY()
On Error GoTo ERR_PERFORM_APOTHECARY
TRIBE_STATUS = "PERFORM_APOTHECARY"
Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
ERR_PERFORM_APOTHECARY_CLOSE:
Exit Function
ERR_PERFORM_APOTHECARY:
Call A999_ERROR_HANDLING
Resume ERR_PERFORM_APOTHECARY_CLOSE
Public Function PERFORM_BAKING()
On Error GoTo ERR_BAKING
TRIBE_STATUS = "PERFORM_BAKING"

TempOutput = "(using "
Dim YEAST As String
   
' NEED TO CHECK FOR BUILDING

Call CHECK_FOR_BUILDING("BAKERY")
   
'BUILDING_FOUND for the building being found or not
'TManufacturingLimit for the total capacity of building/s
'This doesn't check usage so far this month

If BUILDING_FOUND = "N" Then
   If Right(TurnActOutPut, 3) = "^B " Then
      TurnActOutPut = TurnActOutPut & "No bakery found"
   Else
      TurnActOutPut = TurnActOutPut & ", No bakery found"
   End If
   Exit Function
End If
   
Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "BAKER")
   
'Maximum specialists is equal to the smallest of BAKER_FOUND and TSpecialists
If TSpecialists > BAKER_FOUND Then
   TSpecialists = BAKER_FOUND
End If
   
TActives = TActives + (TSpecialists * 2)
   
Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "BAKER", "SPECIALISTS_USED", TSpecialists)

'Reduce the number of Bakers if there is insufficient capacity
If TManufacturingLimit < TActives Then
   TActives = TManufacturingLimit
End If

'Calculate the number of iterations
If TPeople > 0 Then
   TNUMOCCURS = TActives / TPeople
Else
   TNUMOCCURS = 0
End If
'Calculate the number of items made based on the iterations
NumItemsMade = TNUMOCCURS * TNumItems

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "FINE BREAD")
If RESEARCH_FOUND = "Y" Then
    NumItemsMade = (TNUMOCCURS * TNumItems) * 1.5
End If

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "YEAST")
If RESEARCH_FOUND = "Y" Then
   YEAST = "YES"
End If

ModifyTable = "Y"

TCount = 0
Index1 = 1

BRACKET = InStr(TGoods(Index1), "(")
If BRACKET > 0 Then
   ITEM = Left(TGoods(Index1), (BRACKET - 1))
Else
   ITEM = TGoods(Index1)
End If

If Not ITEM = "EMPTY" Then
   If VERIFY_QUANTITY(ITEM, TQuantity(Index1)) = "NO" Then
      Call Calc_New_Num_Occurs
      ' With a new iterations figure, the number of items needs to be recalced
      NumItemsMade = TNUMOCCURS * TNumItems
   End If
End If

If TQuantity(Index1) > 0 Then
   If TQuantity(Index1) = SQuantity(Index1) Then
      TQuantity(Index1) = (TQuantity(Index1) * TNUMOCCURS)
   Else
      TQuantity(Index1) = (SQuantity(Index1) * TNUMOCCURS)
   End If
End If
   
If ModifyTable = "Y" Then
   BRACKET = InStr(TGoods(Index1), "(")
   If BRACKET > 0 Then
      ITEM = Left(TGoods(Index1), (BRACKET - 1))
   Else
      ITEM = TGoods(Index1)
   End If
   If YEAST = "YES" Then
      If ITEM = "FLOUR" Then
         NumItemsMade = NumItemsMade + ((TNUMOCCURS * TNumItems) * 0.5)
      End If
   End If
      
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", TQuantity(Index1))
   
   TempOutput = TempOutput & TQuantity(Index1) & " " & ITEM
       
   BRACKET = InStr(TItem, "(")
   If BRACKET > 0 Then
      ITEM = Left(TItem, (BRACKET - 1))
   Else
      ITEM = TItem
   End If
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "ADD", NumItemsMade)
End If

      
' update output line
If ModifyTable = "Y" Then
   Call UPDATE_TURNACTOUTPUT("NO")
   TempOutput = TempOutput & ")"
   TurnActOutPut = TurnActOutPut & TempOutput
 
End If

ERR_BAKING_CLOSE:
   Exit Function


ERR_BAKING:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_BAKING_CLOSE
End If

End Function

Public Function PERFORM_BLUBBERWORK()
On Error GoTo ERR_PERFORM_BLUBBERWORK
TRIBE_STATUS = "PERFORM_BLUBBERWORK"

Dim People_Used_Cauldron As Long
Dim TCauldrons As Long

TNewWax = 0
TNewOil = 0
Finished = "N"
People_Used_Cauldron = 0
     
TCauldrons = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "CAULDRON")

DoCmd.Hourglass True

Do Until Finished = "Y"
   If TCauldrons > 0 Then
      If TActives > 0 Then
         TActives = TActives - 1
         People_Used_Cauldron = People_Used_Cauldron + 1
         If People_Used_Cauldron = 10 Then
            People_Used_Cauldron = 0
            TCauldrons = TCauldrons - 1
         End If
         AMOUNT_OF_BLUBBER = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BLUBBER")
         If AMOUNT_OF_BLUBBER >= 8 Then
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "SUBTRACT", 8)
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WAX", "ADD", 4)
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "OIL", "ADD", 1)
            TNewWax = TNewWax + 4
            TNewOil = TNewOil + 1
         Else
            Finished = "Y"
         End If
      Else
         Finished = "Y"
      End If
   ElseIf TActives > 0 Then
      TActives = TActives - 1
      AMOUNT_OF_BLUBBER = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BLUBBER")
      If AMOUNT_OF_BLUBBER >= 4 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "SUBTRACT", 4)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WAX", "ADD", 2)
         TNewWax = TNewWax + 2
      Else
         Finished = "Y"
      End If
   Else
      Finished = "Y"
   End If
Loop

If Right(TurnActOutPut, 3) = "^B " Then
   TurnActOutPut = TurnActOutPut & "Bl/wrk (Wax = " & TNewWax & ", Oil = " & TNewOil & ")"
Else
   TurnActOutPut = TurnActOutPut & ", Bl/wrk (Wax = " & TNewWax & ", Oil = " & TNewOil & ")"
End If
   
ERR_PERFORM_BLUBBERWORK_CLOSE:
   Exit Function

ERR_PERFORM_BLUBBERWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_BLUBBERWORK_CLOSE

End Function

Public Function PERFORM_BONEWORK()
On Error GoTo ERR_PERFORM_BONEWORK
TRIBE_STATUS = "PERFORM_BONEWORK"

Dim CHECK_BONE As Long
Dim CHECK_FRAME As Long

Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
Index1 = 1
   If ModifyTable = "Y" Then
      Do Until Index1 > 4
         If TGoods(Index1) = "EMPTY" Then
            Index1 = 4
         Else
            BRACKET = InStr(TGoods(Index1), "(")
            If BRACKET > 0 Then
               ITEM = Left(TGoods(Index1), (BRACKET - 1))
            Else
               ITEM = TGoods(Index1)
            End If
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", (TNUMOCCURS * TQuantity(Index1)))
         End If
         Index1 = Index1 + 1
      Loop
      
      BRACKET = InStr(TItem, "(")
      If BRACKET > 0 Then
         ITEM = Left(TItem, (BRACKET - 1))
      Else
         ITEM = TItem
      End If
      CHECK_BONE = InStr(ITEM, "BONE")
      CHECK_FRAME = InStr(ITEM, "FRAME")
      
      If CHECK_BONE = 0 Then
         If CHECK_FRAME = 0 Then
            ITEM = "BONE " & ITEM
         End If
      End If
      
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "ADD", NumItemsMade)
   End If

   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("NO")
   End If

ERR_PERFORM_BONEWORK_CLOSE:
   Exit Function

ERR_PERFORM_BONEWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_BONEWORK_CLOSE

End Function

Public Function PERFORM_BONING()
On Error GoTo ERR_PERFORM_BONING
TRIBE_STATUS = "PERFORM_BONING"
   
   ' Calc BONING Implements Used.
   
   Call Process_Implement_Usage("BONING", "ALL", TActives, "NO")
       
   Call PERFORM_COMMON("N", "Y", "N", 0, "NONE")

   Index1 = 1
   If ModifyTable = "Y" Then
      If TGoods(Index1) = "GOAT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GOAT", "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 6) * 2))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 6) * 4))
      ElseIf TGoods(Index1) = "CATTLE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CATTLE", "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 3) * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 3) * 20))
      ElseIf TGoods(Index1) = "HORSE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 2) * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 2) * 30))
      ElseIf TGoods(Index1) = "DOG" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "DOG", "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 1) * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 3))
      ElseIf TGoods(Index1) = "CAMEL" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CAMEL", "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 2) * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 2) * 30))
      ElseIf TGoods(Index1) = "ELEPHANT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "ELEPHANT", "SUBTRACT", TQuantity(Index1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", ((TNUMOCCURS * 1) * 12))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 60))
      End If
   End If
      
   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("ASNI")
   End If

ERR_PERFORM_BONING_CLOSE:
   Exit Function

ERR_PERFORM_BONING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_BONING_CLOSE

End Function

Public Function PERFORM_BOOK_WRITING()
On Error GoTo ERR_PERFORM_BOOK_WRITING
TRIBE_STATUS = "PERFORM_BOOK_WRITING"

Dim DLLEVEL As Long
Dim PARCHMENT As Long
Dim CHANCE As Long

COMPRESTAB.index = "TRIBE"
COMPRESTAB.MoveFirst

' Make this a drop down dialog box.

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TEMP_COMP_RESEARCH_LIST;")
qdfCurrent.Execute

Set TEMPCOMPRESTAB = TVDB.OpenRecordset("TEMP_COMP_RESEARCH_LIST")

COMPRESTAB.Seek "=", TTRIBENUMBER

Do
  TEMPCOMPRESTAB.AddNew
  TEMPCOMPRESTAB![CLAN] = TCLANNUMBER
  TEMPCOMPRESTAB![TRIBE] = TTRIBENUMBER
  TEMPCOMPRESTAB![HEXMAP] = Tribes_Current_Hex
  TEMPCOMPRESTAB![TOPIC] = COMPRESTAB![TOPIC]
  TEMPCOMPRESTAB.UPDATE
  COMPRESTAB.MoveNext
  If COMPRESTAB.EOF Then
     Exit Do
  End If
  If Not COMPRESTAB![TRIBE] = TTRIBENUMBER Then
     Exit Do
  End If
  
  If COMPRESTAB.EOF Then
     Exit Do
  End If
 
Loop
  
TEMPCOMPRESTAB.Close

DoCmd.OpenForm "WRITE BOOK", , , , A_EDIT, A_DIALOG

COMPRESTAB.index = "primarykey"

ERR_PERFORM_BOOK_WRITING_CLOSE:
   Exit Function

ERR_PERFORM_BOOK_WRITING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_BOOK_WRITING_CLOSE

End Function

Public Function PERFORM_BRICK_MAKING()
On Error GoTo ERR_PERFORM_BRICK_MAKING
TRIBE_STATUS = "PERFORM_BRICK_MAKING"

' NEED TO CHECK FOR BUILDING

   BUILDING_FOUND = "N"
   
   Call CHECK_FOR_BUILDING("BRICKWORK")

   If BUILDING_FOUND = "N" Then
      TurnActOutPut = TurnActOutPut & "No brickworks found,"
      Exit Function
   End If
   
   Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
    
ERR_PERFORM_BRICK_MAKING_CLOSE:
   Exit Function

ERR_PERFORM_BRICK_MAKING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_BRICK_MAKING_CLOSE

End Function

Public Function PERFORM_CHEESE_MAKING()
On Error GoTo ERR_PERFORM_CHEESE_MAKING
TRIBE_STATUS = "PERFORM_CHEESE_MAKING"
   
   Call PERFORM_COMMON("Y", "Y", "Y", 2, "NONE")

   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("NO")
   End If

ERR_PERFORM_CHEESE_MAKING_CLOSE:
   Exit Function

ERR_PERFORM_CHEESE_MAKING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_CHEESE_MAKING_CLOSE

End Function

Public Function PERFORM_CURING()
On Error GoTo ERR_PERFORM_CURING
TRIBE_STATUS = "PERFORM_CURING"

   Call PERFORM_COMMON("Y", "Y", "N", 3, "NONE")
    
   Index1 = 1
   If ModifyTable = "Y" Then
      Do Until Index1 > 3
         If TGoods(Index1) = "EMPTY" Then
            Index1 = 3
         Else
            BRACKET = InStr(TGoods(Index1), "(")
            If BRACKET > 0 Then
               ITEM = Left(TGoods(Index1), (BRACKET - 1))
            Else
               ITEM = TGoods(Index1)
            End If
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", (TNUMOCCURS * TQuantity(Index1)))
         End If
         Index1 = Index1 + 1
      Loop
      
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LEATHER", "ADD", NumItemsMade)
   End If
      
   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("NO")
   End If

ERR_PERFORM_CURING_CLOSE:
   Exit Function

ERR_PERFORM_CURING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_CURING_CLOSE

End Function

Public Function PERFORM_DISTILLING()
On Error GoTo ERR_PERFORM_DISTILLING
TRIBE_STATUS = "PERFORM_DISTILLING"
Dim RE_ADD_WATER As String

RE_ADD_WATER = "NO"

' ok - rules
' only one type of grog per distillery
' but as many people per grog as there are stills in all distilleries
' get the number of stills

Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "DISTILLER")

If TSpecialists > 0 Then
   If NO_SPECIALISTS_FOUND > TSpecialists Then
      NO_SPECIALISTS_FOUND = TSpecialists
   ElseIf NO_SPECIALISTS_FOUND < TSpecialists Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
   
   Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "DISTILLER", "SPECIALISTS_USED", NO_SPECIALISTS_FOUND)
      
End If

TActives = TActives + (TSpecialists * 2)


' NEED TO CHECK FOR BUILDING

BUILDING_FOUND = "N"
   
Call CHECK_FOR_BUILDING("DISTILLERY")

If BUILDING_FOUND = "N" Then
   TurnActOutPut = TurnActOutPut & "No distillery found,"
   Exit Function
End If
   
' Total Number of Stills across all distilleries * people = TManufacturingLimit
' Number of stills in the next available distillery * people = TBuildingLimit

If TActives > TBuildingLimit Then
   ' using multiple distilleries
   ' Nothing to do
   If TActives > TManufacturingLimit Then
      TActives = TManufacturingLimit
   End If
ElseIf TBuildingLimit < TActives Then
   TActives = TBuildingLimit
End If
   
   
Call DETERMINE_LIQUID_STORAGE
Call DETERMINE_LIQUID_ONHAND
   
'need to look into. ending up with negatives. possibly delete water then do activity and
'read water if have water on hand.
If LIQUID_STORAGE > LIQUID_ONHAND Then
   If CLng((LIQUID_STORAGE - LIQUID_ONHAND) / 20) > TActives Then
      ' do nothing
   Else
      'check for water
      Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WATER")
      If Num_Goods > 0 Then
         ' Delete Water
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WATER", "SUBTRACT", Num_Goods)
         RE_ADD_WATER = "YES"
         TActives = CLng(Num_Goods / 20)
      Else
         TActives = CLng((LIQUID_STORAGE - LIQUID_ONHAND) / 20)
      End If
   End If
Else
   Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WATER")
   If Num_Goods > 0 Then
      ' Delete Water
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WATER", "SUBTRACT", Num_Goods)
      RE_ADD_WATER = "YES"
      TActives = CLng(Num_Goods / 20)
   Else
      TActives = 0
   End If
End If

Call PERFORM_COMMON("Y", "Y", "Y", 2, "ANI")
    
If RE_ADD_WATER = "YES" Then
   TItem = "Water"
   Call PERFORM_GATHERING
End If
   
ERR_PERFORM_DISTILLING_CLOSE:
   Exit Function

ERR_PERFORM_DISTILLING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_DISTILLING_CLOSE

End Function

Public Function PERFORM_DRESSING()
On Error GoTo ERR_PERFORM_DRESSING
TRIBE_STATUS = "PERFORM_DRESSING"

Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
   
Index1 = 1
If ModifyTable = "Y" Then
   Do Until Index1 > 2
      If TGoods(Index1) = "EMPTY" Then
         Index1 = 2
      Else
         BRACKET = InStr(TGoods(Index1), "(")
         If BRACKET > 0 Then
            ITEM = Left(TGoods(Index1), (BRACKET - 1))
         Else
            ITEM = TGoods(Index1)
         End If
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", (TNUMOCCURS * TQuantity(Index1)))
      End If
      Index1 = Index1 + 1
   Loop
   
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LEATHER", "ADD", NumItemsMade)
End If
   
' update output line
If ModifyTable = "Y" Then
   Call UPDATE_TURNACTOUTPUT("NO")
End If

ERR_PERFORM_DRESSING_CLOSE:
   Exit Function

ERR_PERFORM_DRESSING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_DRESSING_CLOSE

End Function

Public Function PERFORM_FISHING()
On Error GoTo ERR_PERFORM_FISHING
TRIBE_STATUS = "PERFORM_FISHING"

Dim TCOASTAL As String
Dim TMOVEMENT As String
Dim TOTAL_SALT As Long
Dim Fish_Imps(20) As String
Dim Fish_Imps_Mods(20) As Long
Dim Fish_Imps_Numbers(20) As Integer
Dim Max_Fishers As Integer

'Variables
' IF IN OCEAN THEN NOT COASTAL ELSE COASTAL
If TRIBES_TERRAIN = "OCEAN" Then
   TCOASTAL = "N"
ElseIf TRIBES_TERRAIN = "LAKE" Then
   TCOASTAL = "N"
Else
   ' search for rivers
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", Tribes_Current_Hex

   If Mid(hexmaptable![Borders], 1, 2) = "RI" Then
      TCOASTAL = "Y"
   End If
   If Mid(hexmaptable![Borders], 3, 2) = "RI" Then
      TCOASTAL = "Y"
   End If
   If Mid(hexmaptable![Borders], 5, 2) = "RI" Then
      TCOASTAL = "Y"
   End If
   If Mid(hexmaptable![Borders], 7, 2) = "RI" Then
      TCOASTAL = "Y"
   End If
   If Mid(hexmaptable![Borders], 9, 2) = "RI" Then
      TCOASTAL = "Y"
   End If
   If Mid(hexmaptable![Borders], 11, 2) = "RI" Then
      TCOASTAL = "Y"
   End If

   ' can't assume this.
   ' must check for ocean in surrounding hexs
   
   N_HEX = GET_MAP_NORTH(Tribes_Current_Hex)
   NE_HEX = GET_MAP_NORTH_EAST(Tribes_Current_Hex)
   SE_HEX = GET_MAP_SOUTH_EAST(Tribes_Current_Hex)
   S_HEX = GET_MAP_SOUTH(Tribes_Current_Hex)
   SW_HEX = GET_MAP_SOUTH_WEST(Tribes_Current_Hex)
   NW_HEX = GET_MAP_NORTH_WEST(Tribes_Current_Hex)
  
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", N_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NE_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SE_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", S_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SW_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NW_HEX

   If Not hexmaptable.NoMatch Then
      If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "OCEAN" Then
         TCOASTAL = "Y"
      End If
   End If

End If
count = 1

Do While count < 21
   Fish_Imps(count) = "EMPTY"
   Fish_Imps_Mods(count) = 0
   Fish_Imps_Numbers(count) = 0
   count = count + 1
Loop

'Number of Actives/Fishermen is held in TActives variable
'Number of Specialists is held in TSpecialists variable
' check availability of specialist.
Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FISHER")
'Number of Implements, includes vessels
'cycle through implement table, pick up the implements
count = 1
ImplementsTable.index = "ACTIVITY"
ImplementsTable.MoveFirst
ImplementsTable.Seek "=", "FISHING", "PROVS"

Do While ImplementsTable![ACTIVITY] = TActivity
   If ImplementsTable.NoMatch Then
      ' Implements not found for fishing - should never happen
   Else
      'Grab Items and modifiers
      ' This will get each of the ships, nets, etc
      
      IMPLEMENT = ImplementsTable![IMPLEMENT]
      IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
   
      ' clan, tribe, activity, item (from activity), implement used
      ' 0330, 0330, fishing, provs, fisher
      PROCESSITEMS.MoveFirst
      PROCESSITEMS.Seek "=", TTRIBENUMBER, TActivity, TItem, IMPLEMENT
      If PROCESSITEMS.NoMatch Then
          ' That Implement was not allocated
      Else
         ' The Implement has been assigned
         
         Fish_Imps(count) = IMPLEMENT
         Fish_Imps_Mods(count) = IMPLEMENT_MODIFIER
         If PROCESSITEMS![QUANTITY] > (TActives + TSpecialists) Then
            Fish_Imps_Numbers(count) = TActives + TSpecialists
         Else
            Fish_Imps_Numbers(count) = PROCESSITEMS![QUANTITY]
         End If
         count = count + 1
      End If
   End If
   ImplementsTable.MoveNext
   If ImplementsTable.EOF Then
      Exit Do
   End If
Loop

' Get number of ships and number of valid fishermen and specialists
count = 1
Max_Fishers = 0

Do While count < 21
   'for each implement see if there is a matching VALID_SHIP
   VALIDSHIPS.MoveFirst
   VALIDSHIPS.Seek "=", Fish_Imps(count)
   If VALIDSHIPS.NoMatch Then
      ' Not a valid ship but still a valid implement
   Else
      ' Valid Ship
      Max_Fishers = Max_Fishers + (VALIDSHIPS![Max_Effect_Fishing] * Fish_Imps_Numbers(count))
   End If
   count = count + 1
   If Fish_Imps(count) = "EMPTY" Then
      count = 21
   End If
Loop

If TCOASTAL = "Y" Then
   'Ignore Max_Fishers
ElseIf Max_Fishers = 0 Then
   ' no boats were specified
ElseIf Max_Fishers >= (TActives + TSpecialists) Then
   ' good to go
ElseIf Max_Fishers <= (TActives + TSpecialists) Then
   ' reduce TSpecialists and TActives
' FOR NOW DO NOT REDUCE FISHERS
'   Do Until Max_Fishers >= (TActives + TSpecialists)
'      If TActives > 0 Then
'         TActives = TActives - 1
'      Else
'         TSpecialists = TSpecialists - 1
'      End If
'   Loop
End If

'Logic Flow
'still need to place modifiers for specialists and actives
'Cycle Implements and calc using modifier
count = 1

Do While count < 21
   TActives = TActives + (Fish_Imps_Mods(count) * Fish_Imps_Numbers(count))
   count = count + 1
   If Fish_Imps(count) = "EMPTY" Then
      count = 21
   End If
Loop

'time to calc number of fish
'Multiply Actives by base figure
TFishing = CLng(TActives * 1.3)
'Multiply Specialists by Actives by base figure
TFishing = TFishing + CLng(TSpecialists * 2.6)

'Multiply fish my skill level and weather
TFishing = TFishing + (CLng((TFishing * FISHING_WEATHER) * ((10 + FISHING_LEVEL) / 10)))

If TFishing <= 0 Then
   TFishing = 1
End If

TFish_Caught = TFishing
If TFish_Caught > 0 Then
   Call Check_Turn_Output(", Caught ", "fish ", "", TFish_Caught, "NO")
End If
    
    
If TDistinction = "SALTING" Then
        TFishing = processSaltingExtraFish(TCLANNUMBER, GOODS_TRIBE, TFishing)
End If

' Update Tribe Table
Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "FISH", "ADD", TFishing)

TSkillok(1) = "N"
' Update output line
DoCmd.Hourglass True



ERR_PERFORM_FISHING_CLOSE:
   Exit Function

ERR_PERFORM_FISHING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FISHING_CLOSE

End Function

Public Function PERFORM_FLENSING()
On Error GoTo ERR_PERFORM_FLENSING
TRIBE_STATUS = "PERFORM_FLENSING"

Dim Num_Whales As Long

Call PERFORM_COMMON("Y", "N", "N", 0, "NONE")
   
ModifyTable = "Y"
   
If ModifyTable = "Y" Then
   TRIBESGOODS.MoveFirst
   If TWhale_Size = "S" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (TActives * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 250)
         Call Check_Turn_Output(", Flensed ", " S/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "M" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (TActives * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 500)
         Call Check_Turn_Output(", Flensed ", " M/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "L" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (TActives * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 750)
         Call Check_Turn_Output(", Flensed ", " L/Whale ", "", 1, "NO")
      End If
   End If
End If
      
ERR_PERFORM_FLENSING_CLOSE:
   Exit Function

ERR_PERFORM_FLENSING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FLENSING_CLOSE

End Function

Public Function PERFORM_FLENSING_AND_PEELING()
On Error GoTo ERR_PERFORM_FLENSING_AND_PEELING
TRIBE_STATUS = "PERFORM_FLENSING_AND_PEELING"

Call PERFORM_COMMON("Y", "N", "N", 0, "NONE")
    
MAXIMUM_ACTIVES_1 = FLENSING_LEVEL * 10
MAXIMUM_ACTIVES_2 = PEELING_LEVEL * 10

ModifyTable = "Y"

If TWhale_Size = "S" Then
   If TActives > 15 Then
      TActives = 15
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If TActives > 10 Then
            FLENSERS = 10
            TActives = TActives - 10
         Else
            FLENSERS = TActives
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If TActives > 5 Then
            PEELERS = 5
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
   End If
ElseIf TWhale_Size = "M" Then
   If TActives > 29 Then
      TActives = 29
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If MAXIMUM_ACTIVES_1 > 20 Then
            If TActives > 20 Then
               FLENSERS = 20
               TActives = TActives - 20
            Else
               FLENSERS = TActives
            End If
         Else
            FLENSERS = MAXIMUM_ACTIVES_1
            TActives = TActives - MAXIMUM_ACTIVES_1
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If TActives > 9 Then
            PEELERS = 9
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
   End If
ElseIf TWhale_Size = "L" Then
   If TActives > 42 Then
      TActives = 42
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If MAXIMUM_ACTIVES_1 > 30 Then
            If TActives > 30 Then
               FLENSERS = 30
               TActives = TActives - 30
            Else
               FLENSERS = TActives
            End If
         Else
            FLENSERS = MAXIMUM_ACTIVES_1
            TActives = TActives - MAXIMUM_ACTIVES_1
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If MAXIMUM_ACTIVES_2 > 12 Then
            If TActives > 12 Then
               PEELERS = 12
            Else
               PEELERS = TActives
            End If
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
   End If
End If
   
If ModifyTable = "Y" Then
   If TWhale_Size = "S" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 250)
         Call Check_Turn_Output(", Flensed & Peeled ", " S/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "M" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 500)
         Call Check_Turn_Output(", Flensed & Peeled ", " M/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "L" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 750)
         Call Check_Turn_Output(", Flensed & Peeled ", " L/Whale ", "", 1, "NO")
      End If
   End If
End If
      
ERR_PERFORM_FLENSING_AND_PEELING_CLOSE:
   Exit Function

ERR_PERFORM_FLENSING_AND_PEELING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FLENSING_AND_PEELING_CLOSE

End Function

Public Function PERFORM_FLENSING_AND_PEELING_AND_BONING()
On Error GoTo ERR_PERFORM_FLENSING_AND_PEELING_AND_BONING
TRIBE_STATUS = "PERFORM_FLENSING_AND_PEELING_AND_BONING"

Call PERFORM_COMMON("Y", "N", "N", 0, "NONE")
    
MAXIMUM_ACTIVES_1 = FLENSING_LEVEL * 10
MAXIMUM_ACTIVES_2 = PEELING_LEVEL * 10
MAXIMUM_ACTIVES_3 = BONING_LEVEL * 10

If TWhale_Size = "S" Then
   If TActives > 19 Then
      TActives = 19
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If TActives > 10 Then
            FLENSERS = 10
            TActives = TActives - 10
         Else
            FLENSERS = TActives
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If TActives > 5 Then
            PEELERS = 5
            TActives = TActives - 5
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
      
      If MAXIMUM_ACTIVES_3 > 0 Then
         If TActives > 4 Then
            BONERS = 4
            TActives = TActives - 4
         Else
            BONERS = TActives
         End If
      Else
         BONERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
      BONERS = 0
   End If
ElseIf TWhale_Size = "M" Then
   If TActives > 35 Then
      TActives = 35
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If MAXIMUM_ACTIVES_1 > 20 Then
            If TActives > 20 Then
               FLENSERS = 20
               TActives = TActives - 20
            Else
               FLENSERS = TActives
            End If
         Else
            FLENSERS = MAXIMUM_ACTIVES_1
            TActives = TActives - MAXIMUM_ACTIVES_1
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If TActives > 9 Then
            PEELERS = 9
            TActives = TActives - 9
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
      
      If MAXIMUM_ACTIVES_3 > 0 Then
         If TActives > 6 Then
            BONERS = 6
            TActives = TActives - 6
         Else
            BONERS = TActives
         End If
      Else
         BONERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
      BONERS = 0
   End If
ElseIf TWhale_Size = "L" Then
   If TActives > 50 Then
      TActives = 50
   End If
   If TActives > 0 Then
      If MAXIMUM_ACTIVES_1 > 0 Then
         If MAXIMUM_ACTIVES_1 > 30 Then
            If TActives > 30 Then
               FLENSERS = 30
               TActives = TActives - 30
            Else
               FLENSERS = TActives
            End If
         Else
            FLENSERS = MAXIMUM_ACTIVES_1
            TActives = TActives - MAXIMUM_ACTIVES_1
         End If
      Else
         FLENSERS = 0
      End If
      
      If MAXIMUM_ACTIVES_2 > 0 Then
         If MAXIMUM_ACTIVES_2 > 12 Then
            If TActives > 12 Then
               PEELERS = 12
            Else
               PEELERS = TActives
            End If
         Else
            PEELERS = TActives
         End If
      Else
         PEELERS = 0
      End If
      
      If MAXIMUM_ACTIVES_3 > 0 Then
         If MAXIMUM_ACTIVES_3 > 8 Then
            If TActives > 8 Then
               BONERS = 8
            Else
               BONERS = TActives
            End If
         Else
            BONERS = TActives
         End If
      Else
         BONERS = 0
      End If
   Else
      FLENSERS = 0
      PEELERS = 0
      BONERS = 0
   End If
End If
   
If ModifyTable = "Y" Then
   If TWhale_Size = "S" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - SMALL", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONE", "ADD", (BONERS * 12))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 250)
         Call Check_Turn_Output(", Flensed,Peeled & Boned ", " S/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "M" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - MEDIUM", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONE", "ADD", (BONERS * 12))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 500)
         Call Check_Turn_Output(", Flensed,Peeled & Boned ", " M/Whale ", "", 1, "NO")
      End If
   ElseIf TWhale_Size = "L" Then
      Num_Whales = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE")
      If Num_Whales >= 1 Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WHALE - LARGE", "SUBTRACT", 1)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BLUBBER", "ADD", (FLENSERS * 10))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (PEELERS * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONE", "ADD", (BONERS * 12))
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 750)
            Call Check_Turn_Output(", Flensed,Peeled & Boned ", " L/Whale ", "", 1, "NO")
         End If
      End If
      
   End If

ERR_PERFORM_FLENSING_AND_PEELING_AND_BONING_CLOSE:
   Exit Function

ERR_PERFORM_FLENSING_AND_PEELING_AND_BONING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FLENSING_AND_PEELING_AND_BONING_CLOSE

End Function

Public Function PERFORM_FLETCHING()
On Error GoTo ERR_PERFORM_FLETCHING
TRIBE_STATUS = "PERFORM_FLETCHING"

   Call PERFORM_COMMON("Y", "Y", "Y", 2, "NONE")
    
ERR_PERFORM_FLETCHING_CLOSE:
   Exit Function

ERR_PERFORM_FLETCHING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FLETCHING_CLOSE

End Function

Public Function PERFORM_FORESTRY()
On Error GoTo ERR_FORESTRY
TRIBE_STATUS = "PERFORM_FORESTRY"

Dim LOGS_USED As Long
Dim COAL_MADE As Long
Dim TOTAL_SCRAPERS As Long
Dim TOTAL_LOGGERS As Long
Dim Initial_Foresters As Long
Dim Initial_Loggers As Long

Initial_Foresters = TActives
Initial_Loggers = TActives

Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FORESTER")
   
If TItem = "BARK" Then
   TBark = 0
   TOTAL_SCRAPERS = TActives
   ' check availability of specialist.
   If TSpecialists > 0 Then
      If TSpecialists > FORESTER_FOUND Then
         TSpecialists = FORESTER_FOUND
      End If
      Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FORESTER", "SPECIALISTS_USED", TSpecialists)
      TOTAL_SCRAPERS = TOTAL_SCRAPERS + TSpecialists
   End If

   Call Process_Implement_Usage("FORESTRY", "BARK", TOTAL_SCRAPERS, "YES")
   
   ' add in the TSpecialists again to achieve the x 2 capability
   TOTAL_SCRAPERS = TOTAL_SCRAPERS + TSpecialists
   
   If TOTAL_SCRAPERS > 0 Then
      TBark = TOTAL_SCRAPERS * 20
   End If
       
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BARK", "ADD", TBark)

   ' Update output line
   If TBark > 0 Then
      Call Check_Turn_Output(",", "effective people stripped ", " bark ", TBark, "YES")
   End If
       
   TSkillok(1) = "N"
       
ElseIf TItem = "LOG" Then
   TLogs = 0
   TOTAL_LOGGERS = TActives
    
   ' check availability of specialist.
   If TSpecialists > 0 Then
      If TSpecialists > FORESTER_FOUND Then
         TSpecialists = FORESTER_FOUND
      End If
      Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FORESTER", "SPECIALISTS_USED", TSpecialists)
      TOTAL_LOGGERS = TOTAL_LOGGERS + TSpecialists
   End If

   ' Calc Implements Used.
   Call Process_Implement_Usage("FORESTRY", "LOG", TOTAL_LOGGERS, "YES")
   
   ' add in the TSpecialists again to achieve the x 2 capability
   TOTAL_LOGGERS = TOTAL_LOGGERS + TSpecialists
   
   If TOTAL_LOGGERS > 0 Then
      TLogs = TOTAL_LOGGERS * LOGS_TO_CUT
      TActives = 0
   End If
          
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LOG", "ADD", TLogs)

   ' Update output line
   If TLogs > 0 Then
      Call Check_Turn_Output(",", "effective people cut ", " logs ", TLogs, "YES")
   End If
      
   TSkillok(1) = "N"
       
ElseIf TItem = "LOG/H" Then
   TLogs = 0
   TOTAL_LOGGERS = TActives
    
   ' check availability of specialist.
   If TSpecialists > 0 Then
      If TSpecialists > FORESTER_FOUND Then
         TSpecialists = FORESTER_FOUND
      End If
      Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FORESTER", "SPECIALISTS_USED", TSpecialists)
      TOTAL_LOGGERS = TOTAL_LOGGERS + TSpecialists
   End If
      
   ' Calc Implements Used.
   Call Process_Implement_Usage("FORESTRY", "LOG", TOTAL_LOGGERS, "YES")
    
   ' add in the TSpecialists again to achieve the x 2 capability
   TOTAL_LOGGERS = TOTAL_LOGGERS + TSpecialists
   
    If TOTAL_LOGGERS > 0 Then
       TLogs = TOTAL_LOGGERS * 2
       TActives = 0
    End If
          
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LOG/H", "ADD", TLogs)

    ' Update output line
    If TLogs > 0 Then
       Call Check_Turn_Output(",", "effective people cut ", " H\Logs ", TLogs, "YES")
    End If
       
    TSkillok(1) = "N"
      
ElseIf TItem = "CHARCOAL MAKING" Then
' NEED TO CHECK FOR BUILDING

  BUILDING_FOUND = "N"
   
  Call CHECK_FOR_BUILDING("CHARHOUSE")

  If BUILDING_FOUND = "N" Then
     TurnActOutPut = TurnActOutPut & "No charhouse found,"
     Exit Function
  End If
  
  RESEARCH_FOUND = "N"

  Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "IMPROVED CHARCOAL MAKING")
        
  If RESEARCH_FOUND = "Y" Then
     TActives = Round(TActives * 1.5, 0)
  End If
  
  ' ALLOW FOR CHARCOAL MAKING
  If TSpecialists > 0 Then
     If TSpecialists > FORESTER_FOUND Then
        TSpecialists = FORESTER_FOUND
     End If
     Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FORESTER", "SPECIALISTS_USED", TSpecialists)
  End If
    
  TActives = TActives + (TSpecialists * 2)

  COMPRESTAB.Seek "=", TTRIBENUMBER, "WOOD CHIPPING"
  If Not COMPRESTAB.NoMatch Then
     LOGS_USED = (TActives * 3)
     COAL_MADE = (TActives * 15)
  Else
     LOGS_USED = (TActives * 2)
     COAL_MADE = (TActives * 10)
  End If
  Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LOG", "SUBTRACT", LOGS_USED)
  Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "COAL", "ADD", COAL_MADE)
  Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "TAR", "ADD", CLng(COAL_MADE / 100))
       
  TurnActOutPut = TurnActOutPut & ", " & TActives & " converted " & LOGS_USED & " logs into " & COAL_MADE & " coal "
End If

ERR_FORESTRY_CLOSE:
   Exit Function

ERR_FORESTRY:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_FORESTRY_CLOSE
End If


End Function

Public Function PERFORM_FURRIER()
On Error GoTo ERR_PERFORM_FURRIER
TRIBE_STATUS = "PERFORM_FURRIER"

Dim TPROVS As Long
Dim TSkins As Long
Dim TFurs As Long

THunters = TActives

' Calc Hunting Implements Used.
Call Process_Implement_Usage("FURRIER", TItem, THunters, "YES")

' Calc REST OF FORMULA
Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FURRIER")
   
If TSpecialists > 0 Then
   If TSpecialists > HUNTER_FOUND Then
      TSpecialists = HUNTER_FOUND
   End If
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FURRIER", "SPECIALISTS_USED", TSpecialists)
End If

THunters = THunters + (TSpecialists * 2)

TProvisions = (THunters * TERRAIN_HUNTING)

TProvisions = CLng((TProvisions * HUNTING_WEATHER) * ((10 + FURRIER_LEVEL) / 10))

If FRESH_WATER > 0 Then
   TProvisions = CLng(TProvisions * FRESH_WATER)
End If

TProvisions = CLng(TProvisions * 0.2)
TPROVS = TProvisions

TSkins = CLng(TPROVS * FURRIER_SKINS)
TFurs = CLng(TPROVS * FURRIER_FURS)

Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", TProvisions)
Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "fur", "ADD", TFurs)
Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "skin", "ADD", TSkins)

TSkillok(1) = "N"
DoCmd.Hourglass True

' Update output line
If TProvisions > 0 Then
   Call Check_Turn_Output(",", " effective people furried (", " Provs ", TProvisions, "YES")
Else
   Call Check_Turn_Output(",", " effective people furried (", " Provs ", 0, "YES")
End If
If TFurs > 0 Then
   Call Check_Turn_Output(",", "", " Furs ", TFurs, "NO")
Else
   Call Check_Turn_Output(",", "", " Furs ", 0, "NO")
End If
If TSkins > 0 Then
   Call Check_Turn_Output(",", "", " Skins ", TSkins, "NO")
Else
   Call Check_Turn_Output(",", "", " Skins ", 0, "NO")
End If
   
   Call Check_Turn_Output(")", "", "", 0, "NO")
   
ERR_PERFORM_FURRIER_CLOSE:
   Exit Function

ERR_PERFORM_FURRIER:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_FURRIER_CLOSE

End Function

Public Function PERFORM_GLASSWORK()
On Error GoTo ERR_PERFORM_GLASSWORK
TRIBE_STATUS = "PERFORM_GLASSWORK"

   Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
    
ERR_PERFORM_GLASSWORK_CLOSE:
   Exit Function

ERR_PERFORM_GLASSWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_GLASSWORK_CLOSE

End Function

Public Function PERFORM_GUT_AND_BONE()
On Error GoTo ERR_PERFORM_GUT_AND_BONE
TRIBE_STATUS = "PERFORM_GUT_AND_BONE"

' Calc Hunting Implements Used.
Call Process_Implement_Usage("BONING", "ALL", TActives, "NO")
       
Call Process_Implement_Usage("GUTTING", "ALL", TActives, "NO")
       
Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
Index1 = 1

If ModifyTable = "Y" Then
   If TGoods(Index1) = "CATTLE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CATTLE", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 4))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 4))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 20))
   ElseIf TGoods(Index1) = "GOAT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GOAT", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 4))
   ElseIf TGoods(Index1) = "HORSE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HORSE", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 30))
   ElseIf TGoods(Index1) = "DOG" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "DOG", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 3))
   ElseIf TGoods(Index1) = "CAMEL" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CAMEL", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 30))
   ElseIf TGoods(Index1) = "ELEPHANT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "ELEPHANT", "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 12))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 12))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 60))
   End If
End If
      
' update output line
If ModifyTable = "Y" Then
   Call UPDATE_TURNACTOUTPUT("ASNI")
End If

ERR_PERFORM_GUT_AND_BONE_CLOSE:
   Exit Function

ERR_PERFORM_GUT_AND_BONE:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_GUT_AND_BONE_CLOSE

End Function

Public Function PERFORM_GUTTING()
On Error GoTo ERR_PERFORM_GUTTING
TRIBE_STATUS = "PERFORM_GUTTING"

' Calc Hunting Implements Used.
   Call Process_Implement_Usage("GUTTING", "ALL", TActives, "NO")

   Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
   Index1 = 1
   If ModifyTable = "Y" Then
      Select Case TGoods(Index1)
      Case "CATTLE", "cattle"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 3)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 3) * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 3) * 20))
      Case "GOAT", "goat"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 6)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 6) * 2))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 6) * 4))
      Case "HORSE", "horse"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 2)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 2) * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 2) * 30))
      Case "DOG", "dog"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 1) * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 3))
      Case "CAMEL", "camel"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 2)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 2) * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 2) * 30))
      Case "ELEPHANT", "elephant"
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", ((TNUMOCCURS * 1) * 12))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 60))
      Case Else
         Call Check_Turn_Output(",", " Animal not catered for in gutting ", "", 0, "YES")
      End Select
      Call UPDATE_TURNACTOUTPUT("ASNI")
   End If

ERR_PERFORM_GUTTING_CLOSE:
   Exit Function

ERR_PERFORM_GUTTING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_GUTTING_CLOSE

End Function

Public Function PERFORM_HARVESTING(CLIMATE As String, CROP As String)
On Error GoTo ERR_HARVESTING
TRIBE_STATUS = "PERFORM_HARVESTING"
       
Dim FARM_TURN As Long
Dim CURR_TURN As Long
Dim Number_Farmers As Long
Dim GET_FARMING_TURN As String
Dim GOOD As String
Dim GOOD_PRODUCED As String
Dim ACRES_HARVESTING As Long
Dim HARVEST_IMPLEMENT As String
Dim CROP_TYPE As String
Dim Num_Trellis As Long

GET_FARMING_TURN = "NO"
MSG1 = "EMPTY"
NO_FARMING = "NO"

' this will allow the harvest to continue is still have actives available
If HARVEST_CONTINUE = "YES" Then
   ' ignore
Else
   HARVEST_CONTINUE = "NO"
End If

WEATHERTABLE.MoveFirst

Set CROP_TABLE = TVDB.OpenRecordset("VALID_CROPS")
CROP_TABLE.index = "PRIMARYKEY"
CROP_TABLE.MoveFirst
CROP_TABLE.Seek "=", CROP

If CROP_TABLE.NoMatch Then
   'MsgBox ("Crop not on Valid_Crops table")
Else
   CROP_TYPE = CROP_TABLE![CROP_TYPE]
   GOOD = CROP_TABLE![GOOD]
   GOOD_PRODUCED = CROP_TABLE![GOOD_PRODUCED]
   ACRES_HARVESTING = CROP_TABLE![ACRES_HARVESTING]
   CROP = CROP_TABLE![CROP]
End If

CROP_TABLE.Close

Set CLIMATETABLE = TVDB.OpenRecordset("VALID_CLIMATE")
CLIMATETABLE.index = "PRIMARYKEY"

CLIMATETABLE.Seek "=", CLIMATE, CROP, "ALL"

If CLIMATETABLE.NoMatch Then
   If InStr(FARMING_TERRAIN, "HILLS") Then
      WEATHER_CROP = CROP & " HILL"
      CLIMATETABLE.Seek "=", CLIMATE, CROP, "HILL"
      If CLIMATETABLE.NoMatch Then
         WEATHER_CROP = CROP
         CLIMATETABLE.Seek "=", CLIMATE, CROP, "MOST"
      End If
   ElseIf InStr(FARMING_TERRAIN, "PRAIRIE") Then
      WEATHER_CROP = CROP & " FLAT"
      CLIMATETABLE.Seek "=", CLIMATE, CROP, "PRAIRIE"
      If CLIMATETABLE.NoMatch Then
         WEATHER_CROP = CROP
         CLIMATETABLE.Seek "=", CLIMATE, CROP, "MOST"
      End If
   Else
      WEATHER_CROP = CROP
      CLIMATETABLE.Seek "=", CLIMATE, CROP, "MOST"
   End If
Else
   WEATHER_CROP = CROP
End If

Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FARMER")

If TSpecialists > 0 Then
   If TSpecialists > FARMER_FOUND Then
      TSpecialists = FARMER_FOUND
   End If
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "FARMER", "SPECIALISTS_USED", TSpecialists)
End If

TActives = TActives + TSpecialists

FARMERS = TActives

' GET BASE FIGURE FROM VALID_CLIMATE
NEW_CROP = CLIMATETABLE![ITEM_NUMBER]
   
If CROP_TYPE = "TEMPORARY" Then
' ==================Harvesting Modification================== (Alex 21.05.2024)
   CURR_TURN = Left(Current_Turn, 2)
   FARMING_TURN1 = "01" & Right(Current_Turn, 4)
   GET_FARMING_TURN = "NO"
   Do
       FarmingTable.MoveFirst
       FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, FARMING_TURN1, CROP
       If (FarmingTable.NoMatch) Then
            GET_FARMING_TURN = "NO"
' Record was not found. Skip this month
       ElseIf FarmingTable![ITEM_NUMBER] > 0 Then
              GET_FARMING_TURN = "YES"
              Exit Do
       End If
            Call GET_FARMING_TURN1
           FARM_TURN = Left(FARMING_TURN1, 2)
           If FARM_TURN > CURR_TURN - 3 Then
               GET_FARMING_TURN = "EARLY"
               Exit Do
           End If
           If FARM_TURN >= 10 Then
               GET_FARMING_TURN = "LATE"
               Exit Do
           End If
   Loop
   
 If (GET_FARMING_TURN = "NO") Then
          Msg = ", NO " & CROP & " to harvest, "
          Call Check_Turn_Output(Msg, " ", "", 0, "NO")
          Exit Function
 End If
If (GET_FARMING_TURN = "LATE") Then
'          Msg = ", Too late to harvest " & CROP & ", "
'          Call Check_Turn_Output(Msg, " ", "", 0, "NO")
          Exit Function
 End If
If (GET_FARMING_TURN = "EARLY") Then
 '         Msg = ", Too early to harvest " & CROP & ", "
 '         Call Check_Turn_Output(Msg, " ", "", 0, "NO")
          Exit Function
 End If
 ' ==================Harvesting Modification END===================================
   GET_FARMING_TURN = "NO"
   GAMES_WEATHER.MoveFirst
   GAMES_WEATHER.Seek "=", CURRENT_WEATHER_ZONE, FARMING_TURN1
   
   CURR_TURN = Left(Current_Turn, 2)
   FARM_TURN = Left(FARMING_TURN1, 2)
   
   If CURR_TURN >= FARM_TURN + 3 Then
      WEATHER_TURN1 = GAMES_WEATHER![WEATHER]
      FarmingTable.Edit
   
      If HARVEST_CONTINUE = "YES" Then
         'implements were added in previous round
      Else
         ' Calc Hunting Implements Used.
         Call Process_Implement_Usage("HARVEST", CROP, FARMERS, "YES")
         ' add in Tspecialists to get the double benefit
         FARMERS = FARMERS + TSpecialists
      End If
      
        'If CROP = "GRAPES" Then
         '   Num_Trellis = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "Trellis")
         '   If Num_Trellis > 0 Then
         '       FARMERS = Round(FARMERS * 1.25, 0)
          '  End If
        'End If
      
      ' Set specialists to zero incase of a loop
      TSpecialists = 0
   
      Do Until FARMERS < 1
         FARMERS = FARMERS - 1
         If FarmingTable![ITEM_NUMBER] >= ACRES_HARVESTING Then
            FarmingTable![ITEM_NUMBER] = FarmingTable![ITEM_NUMBER] - ACRES_HARVESTING
            ACRES_HARVESTED = ACRES_HARVESTED + ACRES_HARVESTING
            'Otherwise extra farmer will be deducted: (AlexD 23.06.2024)
            If FarmingTable![ITEM_NUMBER] = 0 Then Exit Do
           
         Else
            ACRES_HARVESTED = ACRES_HARVESTED + FarmingTable![ITEM_NUMBER]
            FarmingTable![ITEM_NUMBER] = 0
            Exit Do
         End If
      Loop
    
      If FarmingTable![ITEM_NUMBER] < 0 Then
         FarmingTable![ITEM_NUMBER] = 0
      End If

      FarmingTable.UPDATE
      
      NEW_CROP = CLng(NEW_CROP * ACRES_HARVESTED)
      
      'This is actually the place where the crop is modified for turn 1 (first turn of the crop) wheather
      WEATHERTABLE.MoveFirst
      WEATHERTABLE.Seek "=", WEATHER_TURN1, "PLANTING", WEATHER_CROP
      NEW_CROP = NEW_CROP * WEATHERTABLE![Modifier]

      ' MULTIPLY BY WEATHER while growing
      'If Not NO_FARMING = "YES" Then
        Call GET_FARMING_TURN1
        FARM_TURN = Left(FARMING_TURN1, 2)
        Do Until FARM_TURN >= CURR_TURN
            FarmingTable.MoveFirst
            FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, FARMING_TURN1, "WEATHER"
            GAMES_WEATHER.MoveFirst
            GAMES_WEATHER.Seek "=", CURRENT_WEATHER_ZONE, FARMING_TURN1
            WEATHER_TURN2 = GAMES_WEATHER![WEATHER]
            WEATHERTABLE.MoveFirst
            WEATHERTABLE.Seek "=", WEATHER_TURN2, "GROWING", WEATHER_CROP
            NEW_CROP = NEW_CROP * WEATHERTABLE![Modifier]
            Call GET_FARMING_TURN1
            FARM_TURN = Left(FARMING_TURN1, 2)
         Loop
         ' MULTIPLY BY WEATHER while harvesting
         FarmingTable.MoveFirst
         FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, FARMING_TURN1, "WEATHER"
         GAMES_WEATHER.MoveFirst
         GAMES_WEATHER.Seek "=", CURRENT_WEATHER_ZONE, FARMING_TURN1
         WEATHER_TURN3 = GAMES_WEATHER![WEATHER]
         WEATHERTABLE.MoveFirst
         WEATHERTABLE.Seek "=", WEATHER_TURN3, "HARVESTING", WEATHER_CROP
         NEW_CROP = NEW_CROP * WEATHERTABLE![Modifier]
      'End If

      ' FRESH WATER > 0 THEN MULTPILY BY 1.1
      If FRESH_WATER > 0 Then
         NEW_CROP = NEW_CROP * 1.1
      End If

      ' * ((10 + SKILL LEVEL) / 10)
      NEW_CROP = NEW_CROP * ((10 + FARMING_LEVEL) / 10)
 
      'NEW_CROP = CLng(NEW_CROP * ACRES_HARVESTED)
 
      Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Fertiliser")
   
      If RESEARCH_FOUND = "Y" Then
          NEW_CROP = NEW_CROP + CLng(NEW_CROP / 2)
      End If
   
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, GOOD_PRODUCED, "ADD", NEW_CROP)
   
      TOTAL_CROP = TOTAL_CROP + NEW_CROP
   
   Else
      FARMERS = 0
      NO_FARMING = "YES"
      HARVEST_CONTINUE = "NO"
   End If
   
   If FARMERS > 0 Then
      TActives = FARMERS
      FARMERS = 0
      HARVEST_CONTINUE = "YES"
      Call PERFORM_HARVESTING(CLIMATE, CROP)
' Else case added to clean HARVEST_CONTINUE (Alex 21.05.2024)
    Else
      HARVEST_CONTINUE = "NO"
   End If
  
   If MSG1 = "EMPTY" Then
      MSG1 = ", Harvested " & ACRES_HARVESTED & " acres for "
      MSG2 = " " & GOOD_PRODUCED
    
      Call Check_Turn_Output(MSG1, MSG2, "", TOTAL_CROP, "NO")
   End If

   ACRES_HARVESTED = 0
   TOTAL_CROP = 0

Else  ' Permanent
    'harvesting must be performed in turns 7, 8 and/or 9
    'test for month

    Num_Trellis = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "Trellis")
    
    CURR_TURN = Left(Current_Turn, 2)
 
    If ((CURR_TURN >= 7) And (CURR_TURN <= 9)) Then
       ' This area is for harvesting the Permanent Crops.
       PermFarmingTable.MoveFirst
       PermFarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, CROP
       ' Calc Hunting Implements Used.
       Call Process_Implement_Usage("HARVEST", CROP, FARMERS, "YES")
       ' add in additional specialists to get double benefit
       FARMERS = FARMERS + TSpecialists
       If CROP = "GRAPES" Then
          If Num_Trellis > 0 Then
             FARMERS = Round(FARMERS * 1.25, 0)
          End If
       End If
       Do Until FARMERS < 1
          TActives = TActives - 1
          FARMERS = FARMERS - 1
          If (PermFarmingTable![ITEM_NUMBER] - PermFarmingTable![HARVESTED]) >= ACRES_HARVESTING Then
              ACRES_HARVESTED = ACRES_HARVESTED + ACRES_HARVESTING
          Else
              ACRES_HARVESTED = ACRES_HARVESTED + PermFarmingTable![ITEM_NUMBER]
              Exit Do
          End If
       Loop
    
       PermFarmingTable.Edit
       PermFarmingTable![HARVESTED] = PermFarmingTable![HARVESTED] + ACRES_HARVESTED
       PermFarmingTable.UPDATE
      
       ' * ((10 + SKILL LEVEL) / 10)
       NEW_CROP = NEW_CROP * ((10 + FARMING_LEVEL) / 10)
      
       NEW_CROP = CLng(NEW_CROP * ACRES_HARVESTED)
      
       If GOOD_PRODUCED = "HERBS" Then
          Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HERB", "ADD", NEW_CROP)
       Else
          Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, GOOD_PRODUCED, "ADD", NEW_CROP)
       End If
        
       TOTAL_CROP = TOTAL_CROP + NEW_CROP
   
       If MSG1 = "EMPTY" Then
          MSG1 = ", Harvested " & ACRES_HARVESTED & " acres for "
          MSG2 = " " & GOOD_PRODUCED
      
          Call Check_Turn_Output(MSG1, MSG2, "", TOTAL_CROP, "NO")
       End If

       ACRES_HARVESTED = 0
       TOTAL_CROP = 0
    Else
       'WRONG MONTH TO HARVEST PERMANENT CROP
    End If
End If

ERR_HARVESTING_CLOSE:
   Exit Function


ERR_HARVESTING:
If GET_FARMING_TURN = "YES" Then
   If Err = 3021 Then    'NO CURRENT RECORD
      Resume Next
   Else
      Call A999_ERROR_HANDLING
      Resume ERR_HARVESTING_CLOSE
   End If
Else
   Call A999_ERROR_HANDLING
   Msg = TActivity
   MsgBox (Msg)
   Resume ERR_HARVESTING_CLOSE
End If
         
End Function

Public Function PERFORM_HERDING()
On Error GoTo ERR_PERFORM_HERDING
TRIBE_STATUS = "PERFORM_HERDING"

Dim HERDERS_ALLOCATED As String
Dim TOTAL_DOGS As Long
Dim Initial_Animal As Long
Dim Valid_Animals(20) As String
Dim Total_Animals(20) As Long
Dim Animal_Group(20) As Long
Dim ANIMAL As String
Dim Animals_Breed As String
Dim TempAnimals As Single
Dim TNewAnimals As Long
Dim First_Animal_Print As String

count = 1
Do
      Valid_Animals(count) = "EMPTY"
      Total_Animals(count) = 0
      Animal_Group(count) = 0
      count = count + 1
      If count > 20 Then
         Exit Do
      End If
Loop
TNewHorses = 0
TNewLHorses = 0
TNewHHorses = 0
First_Animal_Print = "Yes"

If TItem = "ANIMALS" Then

'output the number of herders and specialists allocated

    If Len(TurnActOutPut) > 20 Then
       HERDERS_ALLOCATED = ", " & TActives & " herders allocated, "
    Else
       HERDERS_ALLOCATED = TActives & " herders allocated, "
    End If
   
    If TSpecialists > 0 Then
        HERDERS_ALLOCATED = HERDERS_ALLOCATED & TSpecialists & " specialists allocated, "
    End If
  
    Call Check_Turn_Output(HERDERS_ALLOCATED, "", "", 0, "NO")
  
    Call HERDING_LIMIT

    Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "HERDER", "SPECIALISTS_USED", TSpecialists)
    
    Call Check_Turn_Output("", "", "", 0, "NO")
    
    TActives = TActives + NO_SPECIALISTS_FOUND
    
    If HERDERSREQ <= TActives Then
        'GREAT'
    Else
        'LOSE HERD
    End If
   
    Set Goods_Tribes_Processed = TVDBGM.OpenRecordset("Goods_Tribes_Processing")
    Goods_Tribes_Processed.index = "primarykey"
    Goods_Tribes_Processed.MoveFirst
    Goods_Tribes_Processed.Seek "=", GOODS_TRIBE
   
    If Goods_Tribes_Processed.NoMatch Then
       Goods_Tribes_Processed.AddNew
       Goods_Tribes_Processed![GOODS_TRIBE] = GOODS_TRIBE
       Goods_Tribes_Processed![Herd_Processed] = "Y"
       Goods_Tribes_Processed.UPDATE
    ElseIf Goods_Tribes_Processed![Herd_Processed] = "Y" Then
       Call Check_Turn_Output("", "Breeding already performed", "", 0, "NO")
       GoTo END_HERDING_CLOSE
    Else
       Goods_Tribes_Processed.Edit
       Goods_Tribes_Processed![Herd_Processed] = "Y"
       Goods_Tribes_Processed.UPDATE
    End If
   
   ' define an array, populate it, loop through it
   ' read valid animals and populate array
   
   VALIDANIMALS.MoveFirst
   ANIMAL = VALIDANIMALS![ANIMAL]
   Do
        count = 1
        Initial_Animal = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ANIMAL)
        If Initial_Animal > 0 Then
            Do
                If Valid_Animals(count) = "EMPTY" Then
                    Valid_Animals(count) = VALIDANIMALS![ANIMAL]
                    Valid_Animals(count) = VALIDANIMALS![UPDATE_ANIMAL]
                    Total_Animals(count) = Initial_Animal
                    Animal_Group(count) = VALIDANIMALS![HERDING_GROUP]
                    Exit Do
                End If
                If Valid_Animals(count) = VALIDANIMALS![UPDATE_ANIMAL] Then
                    Total_Animals(count) = Total_Animals(count) + Initial_Animal
                    Exit Do
                End If
                count = count + 1
                If count > 20 Then
                    Exit Do
                End If
           Loop
        End If
        VALIDANIMALS.MoveNext
        If VALIDANIMALS.EOF Then
            Exit Do
        End If
        ANIMAL = VALIDANIMALS![ANIMAL]
   Loop
   
   Animals_Breed = "NO"

   count = 1
   Do
         TempAnimals = 0
         SWAPSINEFFECT = 0
         SWAPSINEFFECT = HERD_SWAPS(TCLANNUMBER, TTRIBENUMBER, "AAA", "AAA", Valid_Animals(count), "N")
     
         If Total_Animals(count) > 0 Then
             If Animals_Breed = "NO" Then
                 Call Check_Turn_Output("", "Bred (", "", 0, "NO")
                 Animals_Breed = "YES"
             End If
             If Animal_Group(count) = 1 Then
                 TempAnimals = (TERRAIN_HERDING_GROUP_1 * WEATHER_HERDING_GROUP_1)
             ElseIf Animal_Group(count) = 2 Then
                 TempAnimals = (TERRAIN_HERDING_GROUP_2 * WEATHER_HERDING_GROUP_2)
             ElseIf Animal_Group(count) = 3 Then
                 TempAnimals = (TERRAIN_HERDING_GROUP_3 * WEATHER_HERDING_GROUP_3)
             End If
             TempAnimals = ((TempAnimals * ((10 + HERDING_LEVEL) / 10)) / 100)
             TempAnimals = (TempAnimals * (((SWAPSINEFFECT * 10) + 100) / 100))
             TNewAnimals = CLng(TempAnimals * Total_Animals(count))
             If Valid_Animals(count) = "Horse" Then
                 TNewHorses = TNewAnimals
                 TNewAnimals = 0
                  Exit Do
             ElseIf Valid_Animals(count) = "Horse/Light" Then
                 TNewLHorses = TNewAnimals
                 TNewAnimals = 0
                 Exit Do
             ElseIf Valid_Animals(count) = "Horse/Heavy" Then
                 TNewHHorses = TNewAnimals
                 TNewAnimals = 0
                 Exit Do
             End If
             If TNewAnimals > 0 Then
                If First_Animal_Print = "Yes" Then
                   Call Check_Turn_Output("", Valid_Animals(count), "", TNewAnimals, "NO")
                   First_Animal_Print = "No"
               Else
                   Call Check_Turn_Output(", ", Valid_Animals(count), "", TNewAnimals, "NO")
                End If
             End If
             Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, Valid_Animals(count), "ADD", TNewAnimals)
         End If
         count = count + 1
         If count > 20 Then
             Exit Do
         End If
    Loop
   
    If TNewHorses > 0 Then
         ' IF GOT RESEARCH 'Horse Light' THEN 10% OF GROWTH GOES TO LIGHT HORSE
        RESEARCH_FOUND = "N"
        Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Horse Light")
        If RESEARCH_FOUND = "Y" Then
            TNewLHorses = TNewLHorses + CLng(TNewHorses / 10)
            TNewHorses = TNewHorses - CLng(TNewHorses / 10)
        End If
        ' IF GOT RESEARCH 'Horse Heavy' THEN 10% OF GROWTH GOES TO HEAVY HORSE
        RESEARCH_FOUND = "N"
        Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Horse Heavy")
        If RESEARCH_FOUND = "Y" Then
            TNewHHorses = TNewHHorses + CLng(TNewHorses / 10)
            TNewHorses = TNewHorses - CLng(TNewHorses / 10)
        End If
        Call Check_Turn_Output(", ", " horses ", "", TNewHorses, "NO")
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HORSE", "ADD", TNewHorses)
        ' UPDATE WITH HORSE GROWTH SO THAT THEY CAN BE CONVERTED TO WARHORSES IF POSSIBLE
        
        If Turn_Info_Req_NxTurn.BOF Then
            ' do nothing
        Else
            Turn_Info_Req_NxTurn.MoveFirst
        End If
        Turn_Info_Req_NxTurn.Seek "=", TCLANNUMBER, GOODS_TRIBE, "Horses Bred"
   
        If Turn_Info_Req_NxTurn.NoMatch Then
            Turn_Info_Req_NxTurn.AddNew
            Turn_Info_Req_NxTurn![CLAN] = TCLANNUMBER
            Turn_Info_Req_NxTurn![TRIBE] = GOODS_TRIBE
            Turn_Info_Req_NxTurn![ITEM] = "Horses Bred"
            Turn_Info_Req_NxTurn![ITEM_NUMBER] = TNewHorses
            Turn_Info_Req_NxTurn.UPDATE
        Else
            Turn_Info_Req_NxTurn.Edit
            Turn_Info_Req_NxTurn![ITEM_NUMBER] = TNewHorses
            Turn_Info_Req_NxTurn.UPDATE
        End If
    End If
    If TNewLHorses > 0 Then
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HORSE/LIGHT", "ADD", TNewLHorses)
    End If
    If TNewHHorses > 0 Then
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HORSE/HEAVY", "ADD", TNewHHorses)
    End If
   
    TSkillok(1) = "N"
    DoCmd.Hourglass True

   ' Update output line
   If Animals_Breed = "NO" Then
      Call Check_Turn_Output("", " No Breeding Performed", "", 0, "NO")
   Else
      Call Check_Turn_Output("", ") ", "", 0, "NO")
   End If


Else ' today defaults to milking
   Initial_Animal = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "CATTLE")
   
   If Len(TTRIBENUMBER) > 4 Then
      SKILLSTABLE.Seek "=", Left(TTRIBENUMBER, 4), TItem
   Else
      SKILLSTABLE.Seek "=", TTRIBENUMBER, TItem
   End If
   
   If TActives > (HERDING_LEVEL * 10) Then
      TActives = (HERDING_LEVEL * 10)
   End If
   
   ' check availability of specialist.
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "HERDER")
      
   SPECIALIST_FOUND = "Y"

   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If

   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "HERDER", "SPECIALISTS_USED", TSpecialists)
   
   TActives = TActives + (TSpecialists * 2)
   
   If Initial_Animal >= (TActives * 10) Then
      TOTALMILK = (TActives * 10) * 10
   Else
      TOTALMILK = Initial_Animal * 10
   End If
             
   If TOTALMILK > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "MILK", "ADD", TOTALMILK)
   End If
   
   TurnActOutPut = TurnActOutPut & TOTALMILK & " milk, "
   TSkillok(1) = "N"

End If

END_HERDING_CLOSE:
Exit Function

ERR_PERFORM_HERDING:
If (Err = 3021) Then
   
   Resume Next
  
Else
   Call A999_ERROR_HANDLING
   Msg = TActivity
   MsgBox (Msg)
   Resume END_HERDING_CLOSE

End If

End Function

Public Function PERFORM_HUNTING()
On Error GoTo ERR_HUNTING
TRIBE_STATUS = "PERFORM_HUNTING"

Dim Mongol_Hunt As String
Dim Mongol_Hunt2 As String
Dim HUNTERS_ALLOCATED As String
Dim Hunting_Dogs_Available As Long

Mongol_Hunt = "N"
Mongol_Hunt2 = "N"

TActivesAtStart = TActives

' check availability of specialist.
Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "HUNTER")
   
If TSpecialists > HUNTER_FOUND Then
   TSpecialists = HUNTER_FOUND
End If
          
Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "HUNTER", "SPECIALISTS_USED", TSpecialists)

TActives = TActives + TSpecialists
THunters = TActives

' Calc Hunting Implements Used.
Call Process_Implement_Usage("HUNTING", "ALL", TActives, "NO")


' Research Specific Benefits
' Mongol Hunts & Mongol Hunts2

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Mongol Hunts")
        
If RESEARCH_FOUND = "Y" Then
    Mongol_Hunt = "Y"
Else
    Mongol_Hunt = "N"
End If

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Mongol Hunts2")
        
If RESEARCH_FOUND = "Y" Then
    Mongol_Hunt2 = "Y"
Else
    Mongol_Hunt2 = "N"
End If



If Mongol_Hunt2 = "Y" Then
   TActives = TActives + (TActivesAtStart * 0.4)
ElseIf Mongol_Hunt = "Y" Then
   TActives = TActives + (TActivesAtStart * 0.2)
End If

' add TSpecialist again to ensure specialist double benefit
TActives = TActives + TSpecialists

'Add Hunting Dogs
RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "HUNTING DOGS")
        
If RESEARCH_FOUND = "Y" Then
   Hunting_Dogs_Available = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "HUNTING DOG")
Else
   Hunting_Dogs_Available = 0
End If

TActives = TActives + Hunting_Dogs_Available

' Calc REST OF FORMULA
TProvisions = (TActives * TERRAIN_HUNTING)

TProvisions = CLng((TProvisions * HUNTING_WEATHER) * ((10 + HUNTING_LEVEL) / 10))
       
If FRESH_WATER > 0 Then
   TProvisions = CLng(TProvisions * FRESH_WATER)
End If

If ROAMING_HERD = "Y" Then
   TProvisions = CLng(TProvisions * 1.1)
End If

Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", TProvisions)

TSkillok(1) = "N"
DoCmd.Hourglass True

' Update output line
If TProvisions > 0 Then
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output(" ", " effective people hunted", " provs", TProvisions, "YES")
   Else
      Call Check_Turn_Output(", ", " effective people hunted", " provs", TProvisions, "YES")
   End If
End If

END_HUNTING:
Exit Function

ERR_HUNTING:
If (Err = 3021) Then

   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_HUNTING

End If

End Function

Public Function PERFORM_JEWELLERY()
On Error GoTo ERR_PERFORM_JEWELLERY
TRIBE_STATUS = "PERFORM_JEWELLERY"

    Call PERFORM_COMMON("Y", "Y", "Y", 4, "NONE")
    
ERR_PERFORM_JEWELLERY_CLOSE:
   Exit Function

ERR_PERFORM_JEWELLERY:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_JEWELLERY_CLOSE

End Function

Public Function PERFORM_KILLING()
On Error GoTo ERR_PERFORM_KILLING
TRIBE_STATUS = "PERFORM_KILLING"

If TItem = "GOAT" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GOAT", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 4))
ElseIf TItem = "CATTLE" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CATTLE", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 20))
ElseIf TItem = "HORSE" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HORSE", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 30))
ElseIf TItem = "DOG" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HERDING DOG", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 3))
ElseIf TItem = "CAMEL" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "CAMEL", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 30))
ElseIf TItem = "ELEPHANT" Then
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "ELEPHANT", "SUBTRACT", TActives)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 60))
ElseIf TItem = "PEOPLE" Then
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   If TDistinction = "WARRIORS" Then
      TRIBESINFO!INACTIVES = TRIBESINFO!WARRIORS - TActives
      TRIBESINFO.UPDATE
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 4))
   ElseIf TDistinction = "ACTIVES" Then
      TRIBESINFO!INACTIVES = TRIBESINFO!ACTIVES - TActives
      TRIBESINFO.UPDATE
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 4))
   ElseIf TDistinction = "INACTIVES" Then
      TRIBESINFO!INACTIVES = TRIBESINFO!INACTIVES - TActives
      TRIBESINFO.UPDATE
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Provs", "ADD", (TActives * 4))
   End If
End If

NumItemsMade = TActives

Call UPDATE_TURNACTOUTPUT("NO")

ERR_PERFORM_KILLING_CLOSE:
   Exit Function

ERR_PERFORM_KILLING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_KILLING_CLOSE

End Function

Public Function PERFORM_LEATHERWORK()
On Error GoTo ERR_PERFORM_LEATHERWORK
TRIBE_STATUS = "PERFORM_LEATHERWORK"

   Call PERFORM_COMMON("Y", "Y", "Y", 3, "ANI")
    
ERR_PERFORM_LEATHERWORK_CLOSE:
   Exit Function

ERR_PERFORM_LEATHERWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_LEATHERWORK_CLOSE

End Function

Public Function PERFORM_METALWORK()
On Error GoTo ERR_METALWORK
TRIBE_STATUS = "PERFORM_METALWORK"
   
   Initial_Metalworkers = TActives
   TMetalworkers = 0
  
   Call Process_Implement_Usage("METALWORK", TItem, TMetalworkers, "NO")
       
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "METALWORKER")
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
          
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "METALWORKER", "SPECIALISTS_USED", TSpecialists)
 
   TActives = TActives + (TSpecialists * 2) + TMetalworkers

   Call PERFORM_COMMON("Y", "Y", "Y", 5, "ANI")
    
ERR_METALWORK_CLOSE:
Exit Function

ERR_METALWORK:
If (Err = 3021) Then
   
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_METALWORK_CLOSE

End If
    
End Function

Public Function PERFORM_MILLING()
On Error GoTo ERR_MILLING
TRIBE_STATUS = "PERFORM_MILLING"

' NEED TO CHECK FOR BUILDING

   BUILDING_FOUND = "N"
   
   Call CHECK_FOR_BUILDING("MILL")

   If BUILDING_FOUND = "N" Then
      TurnActOutPut = TurnActOutPut & "No mill found,"
      Exit Function
   End If

   ' check availability of specialist.
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "MILLER")
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
          
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "MILLER", "SPECIALISTS_USED", TSpecialists)
 
   TActives = TActives + (TSpecialists * 2)
   
   Call PERFORM_COMMON("Y", "Y", "Y", 1, "NONE")
    
ERR_MILLING_CLOSE:
   Exit Function


ERR_MILLING:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_MILLING_CLOSE
End If

End Function

Public Function PERFORM_MINING()
On Error GoTo ERR_MINING
TRIBE_STATUS = "PERFORM_MINING"

Dim SLAVES_USED As Long
Dim APP_TOOL As String
Dim NEWDIRECTION As String
Dim NEWHEX As String
Dim MINERAL As String
Dim MINERAL_TO_MINE As String
Dim TInitialMiners As Long


RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "APROPRIATE MINING TOOL")
        
If RESEARCH_FOUND = "Y" Then
   APP_TOOL = "YES"
Else
   APP_TOOL = "NO"
   RESEARCH_FOUND = "N"
   Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "APPROPRIATE MINING TOOL")

   If RESEARCH_FOUND = "Y" Then
      APP_TOOL = "YES"
   Else
      APP_TOOL = "NO"
   End If
End If

HEXMAPMINERALS.MoveFirst
HEXMAPMINERALS.Seek "=", Tribes_Current_Hex
       
TAccident = (15 - (Skill_Level_1 * 2) + MINING_ACCIDENTS_WEATHER)
If Len(TTRIBENUMBER) > 4 Then
   DICE_TRIBE = Left(TTRIBENUMBER, 4)
ElseIf Left(TTRIBENUMBER, 1) = "B" Then
   DICE_TRIBE = CLng(TCLANNUMBER)
ElseIf Left(TTRIBENUMBER, 1) = "M" Then
   DICE_TRIBE = CLng(TCLANNUMBER)
Else
   DICE_TRIBE = CLng(TTRIBENUMBER)
End If
DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 1, 0)
DICE2 = DROLL(6, 1, 20, 0, DICE_TRIBE, 1, 0)

If DICE1 <= TAccident Then
   TLossMiners = CLng((((5 + DICE2) - (2 * MINING_LEVEL)) / 100) * TActives)

   If TLossMiners <= 0 Then
      TLossMiners = 0
   Else
      Msg = "Tribe " & TTRIBENUMBER & "Miners Lost = " & TLossMiners
      MsgBox (Msg)
   End If

   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER

   If TLossMiners > 0 Then
      TurnActOutPut = TurnActOutPut & TLossMiners & " Miners Died "
   End If

   Msg = "How many slaves of " & TTRIBENUMBER & " where used in mining?"
   SLAVES_USED = InputBox(Msg, "MINING_LOSS", "0")
   If SLAVES_USED > 0 Then
      If TLossMiners > SLAVES_USED Then
         TRIBESINFO.Edit
         TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - SLAVES_USED
         TRIBESINFO.UPDATE
         TLossMiners = TLossMiners - SLAVES_USED
      Else
         TRIBESINFO.Edit
         TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - TLossMiners
         TRIBESINFO.UPDATE
         TLossMiners = 0
      End If
   End If

   If TLossMiners > 0 Then
      TRIBESINFO.Edit
      TRIBESINFO![MORALE] = TRIBESINFO![MORALE] - 0.01
      MORALELOSS = "Y"
      TRIBESINFO.UPDATE
   End If

   Do Until TLossMiners < 1
      TRIBESINFO.Edit
      If TRIBESINFO![INACTIVES] > TRIBESINFO![ACTIVES] Then
         If TRIBESINFO![INACTIVES] > TRIBESINFO![WARRIORS] Then
            TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] - 1
            TLossMiners = TLossMiners - 1
         Else
            TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - 1
            TLossMiners = TLossMiners - 1
         End If
      ElseIf TRIBESINFO![ACTIVES] > TRIBESINFO![WARRIORS] Then
             TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] - 1
             TLossMiners = TLossMiners - 1
      ElseIf TRIBESINFO![WARRIORS] > TRIBESINFO![INACTIVES] Then
             TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - 1
             TLossMiners = TLossMiners - 1
      ElseIf TRIBESINFO![WARRIORS] = TRIBESINFO![ACTIVES] Then
             If TRIBESINFO![WARRIORS] = TRIBESINFO![ACTIVES] Then
                TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - 1
                TLossMiners = TLossMiners - 1
             Else
                TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] - 1
                TLossMiners = TLossMiners - 1
             End If
      ElseIf TRIBESINFO![ACTIVES] = TRIBESINFO![INACTIVES] Then
             TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] - 1
             TLossMiners = TLossMiners - 1
      Else
             TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] - 1
             TLossMiners = TLossMiners - 1
      End If
      TRIBESINFO.UPDATE
   Loop

End If

If TItem = "LOW YIELD EXTRACTION" Then
   ' EXPAND THIS TO DEAL WITH MULTIPLE ORE TYPES IN SAME HEX
   NEWDIRECTION = InputBox("Which direction are we to mine in?", "MINING", "0")
   If NEWDIRECTION = "N" Then
      NEWHEX = MAP_N
   ElseIf NEWDIRECTION = "NE " Then
      NEWHEX = MAP_NE
   ElseIf NEWDIRECTION = "SE " Then
      NEWHEX = MAP_SE
   ElseIf NEWDIRECTION = "S " Then
      NEWHEX = MAP_S
   ElseIf NEWDIRECTION = "SW " Then
      NEWHEX = MAP_SW
   ElseIf NEWDIRECTION = "NW " Then
      NEWHEX = MAP_NW
   End If
   
   HEXMAPMINERALS.Seek "=", NEWHEX

   If Not IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
      If Not HEXMAPMINERALS![ORE_TYPE] = "NONE" Then
         MINERAL = "YES"
         TItem = HEXMAPMINERALS![ORE_TYPE]
      Else
         MINERAL = "NO"
         TurnActOutPut = TurnActOutPut & ", Mined Nil "
      End If
   Else
      MINERAL = "NO"
      TurnActOutPut = TurnActOutPut & ", Mined Nil "
   End If
ElseIf TItem = "SURVEYING" Then
   ' EXPAND THIS TO DEAL WITH MULTIPLE ORE TYPES IN SAME HEX
   If Not HEXMAPMINERALS.NoMatch Then
      If Not IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
         If Not HEXMAPMINERALS![ORE_TYPE] = "NONE" Then
            MINERAL = "YES"
            TItem = HEXMAPMINERALS![ORE_TYPE]
         Else
            MINERAL = "NO"
            TurnActOutPut = TurnActOutPut & ", Mined Nil "
         End If
      Else
         MINERAL = "NO"
         TurnActOutPut = TurnActOutPut & ", Mined Nil "
      End If
   Else
      MINERAL = "NO"
      TurnActOutPut = TurnActOutPut & ", Mined Nil "
   End If
Else
   ' EXPAND THIS TO DEAL WITH MULTIPLE ORE TYPES IN SAME HEX
   If Not HEXMAPMINERALS.NoMatch Then
       RESEARCH_FOUND = "N"
       Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Geologists")
       If RESEARCH_FOUND = "Y" Then
           If TItem = HEXMAPMINERALS![ORE_TYPE] Then
               MINERAL = "YES"
           ElseIf TItem = HEXMAPMINERALS![SECOND_ORE] Then
               MINERAL = "YES"
           ElseIf TItem = HEXMAPMINERALS![THIRD_ORE] Then
               MINERAL = "YES"
           ElseIf TItem = HEXMAPMINERALS![FORTH_ORE] Then
               MINERAL = "YES"
           Else
               MINERAL = "NO"
           End If
      ElseIf Not IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
         If Not HEXMAPMINERALS![ORE_TYPE] = "NONE" Then
            MINERAL = "YES"
            TItem = HEXMAPMINERALS![ORE_TYPE]
         Else
            MINERAL = "NO"
            TurnActOutPut = TurnActOutPut & ", Mined Nil "
         End If
      Else
         MINERAL = "NO"
         TurnActOutPut = TurnActOutPut & ", Mined Nil "
      End If
   Else
         If TRIBES_TERRAIN = "LOW VOLCANO MOUNTAINs" Then
             MINERAL = "YES"
             TItem = "SULPHUR"
         Else
               MINERAL = "NO"
               TurnActOutPut = TurnActOutPut & ", Mined Nil "
         End If
   End If
End If

If MINERAL = "YES" Then
   TActivesAvailable = TActivesAvailable - TActives
     
   ' check availability of specialist.
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "MINER")
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
          
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "MINER", "SPECIALISTS_USED", TSpecialists)
 
   TActives = TActives + TSpecialists
   TMiners = TActives
   TInitialMiners = TActives
   
   Call Process_Implement_Usage(TActivity, TItem, TMiners, "NO")

   ' Allow for double benefit of specialists
   TMiners = TMiners + TSpecialists
   
'  going to have to fix this at some point.
'  needs to be automated
   ImplementUsage.MoveFirst
   ImplementUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, "LANTERN"

   If Not ImplementUsage.NoMatch Then
      TImplement = InputBox("How many LANTERNS used for mining?", "MINING", "0")
      If TImplement > 0 Then
         total_available = ImplementUsage![total_available] - ImplementUsage![Number_Used]
         If TImplement > total_available Then
            TImplement = total_available
         End If
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "OIL"
         If Not TRIBESGOODS.NoMatch Then
            TRIBESGOODS.Edit
            If TRIBESGOODS![ITEM_NUMBER] >= TImplement Then
               TMiners = TMiners + (TImplement * 0.5)
               TActives = 0
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TImplement
               TRIBESGOODS.UPDATE
               ImplementUsage.Edit
               ImplementUsage![Number_Used] = ImplementUsage![Number_Used] + TImplement
               ImplementUsage.UPDATE
            Else
               TMiners = TMiners + (TRIBESGOODS![ITEM_NUMBER] * 0.5)
               TActives = 0
               TRIBESGOODS![ITEM_NUMBER] = 0
               TRIBESGOODS.UPDATE
               ImplementUsage.Edit
               ImplementUsage![Number_Used] = ImplementUsage![Number_Used] + TImplement
              ImplementUsage.UPDATE
            End If
         End If
      End If
   End If

   ' CHECK FOR MINESHAFTS - MODIFIER IS 50% THEN
   'Mineshafts 4 men per 1 yard, requires 1 log
               '1 miner per 1 yard -= 2 x's output

   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, TCLANNUMBER, "MINESHAFT"

   If Not HEXMAPCONST.NoMatch Then
      If TInitialMiners > HEXMAPCONST![1] Then
         TMiners = TMiners + HEXMAPCONST![1]
      Else
         TMiners = TMiners + TInitialMiners
      End If
   End If

   RESEARCH_FOUND = "N"
   Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Drift Mining")
   If RESEARCH_FOUND = "Y" Then
       TMiners = TMiners + CLng(TInitialMiners / 2)
   End If
  
   VALIDMINERALS.MoveFirst
   VALIDMINERALS.Seek "=", TItem
   If VALIDMINERALS.NoMatch Then
        MsgBox "No match found in VALIDMINERALS table for Tribe: " & TTRIBENUMBER & ", Item: " & TItem, vbExclamation, "No Match Error"
    End If
   ' Calc REST OF FORMULA
   TNEWORE = CLng(TMiners * ((MINING_LEVEL + 2) / 4))
   TNEWORE = CLng(TNEWORE * VALIDMINERALS![MINING_VALUE_1])
   TNEWORE = CLng(TNEWORE * MINING_WEATHER)
   If TItem = "LOW YIELD EXTRACTION" Then
      TNEWORE = CLng(TNEWORE / 2)
   End If
     
   If TNEWORE = 0 Then
      TNEWORE = 1
   End If

   ' Update Tribe Table
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TItem, "ADD", TNEWORE)
   DoCmd.Hourglass True
   
   ' Update output line
   If TNEWORE > 0 Then
      TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people mined " & TNEWORE & " " & TItem
   End If
   TSkillok(1) = "N"
End If

ERR_MINING_CLOSE:
   Exit Function


ERR_MINING:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_MINING_CLOSE
End If


End Function

Public Function PERFORM_MUSIC()
On Error GoTo ERR_PERFORM_MUSIC
TRIBE_STATUS = "PERFORM_MUSIC"

   Call PERFORM_COMMON("Y", "Y", "Y", 5, "NONE")
    
ERR_PERFORM_MUSIC_CLOSE:
   Exit Function

ERR_PERFORM_MUSIC:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_MUSIC_CLOSE

End Function

Public Function PERFORM_PEELING()
On Error GoTo ERR_PERFORM_PEELING
TRIBE_STATUS = "PERFORM_PEELING"

   Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
   Whale = InputBox("What size WHALE is being Peeled? (S/M/L)", "Size", "N")
              
   If ModifyTable = "Y" Then
      TRIBESGOODS.MoveFirst
      If Whale = "S" Then
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "WHALE - SMALL"
         If TRIBESGOODS![ITEM_NUMBER] >= 1 Then
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
            TRIBESGOODS.UPDATE
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (TActives * 10))
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 250)
         End If
      ElseIf Whale = "M" Then
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "WHALE - MEDIUM"
         If TRIBESGOODS![ITEM_NUMBER] >= 1 Then
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
            TRIBESGOODS.UPDATE
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (TActives * 10))
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 500)
         End If
      ElseIf Whale = "L" Then
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "WHALE - LARGE"
         If TRIBESGOODS![ITEM_NUMBER] >= 1 Then
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
            TRIBESGOODS.UPDATE
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (TActives * 10))
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", 750)
         End If
      End If
      
   End If
      
   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("NO")
   End If

ERR_PERFORM_PEELING_CLOSE:
   Exit Function

ERR_PERFORM_PEELING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_PEELING_CLOSE

End Function

Public Function PERFORM_POTTERY()
On Error GoTo ERR_PERFORM_POTTERY
TRIBE_STATUS = "PERFORM_POTTERY"

If TItem = "CLAY" Then
   ' NEED TO CHECK TERRAIN OR SURROUNDING HEX'S - MUST BE NEAR RIVER/LAKE, OR
   ' HAVE STREAMS, SPRINGS IN HEX
    
   ACTIVES_INUSE = TActives
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "CLAY"
        
   Call Process_Implement_Usage(TActivity, TItem, TActives, "NO")

   TOTAL_CLAY = (TActives * 20)
   
   If TRIBESGOODS.NoMatch Then
      TRIBESGOODS.AddNew
      TRIBESGOODS![CLAN] = TCLANNUMBER
      TRIBESGOODS![TRIBE] = GOODS_TRIBE
      TRIBESGOODS![ITEM_TYPE] = "RAW"
      TRIBESGOODS![ITEM] = "CLAY"
      TRIBESGOODS![ITEM_NUMBER] = TOTAL_CLAY
   Else
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + TOTAL_CLAY
   End If
 
   TRIBESGOODS.UPDATE
        
   Call Check_Turn_Output(", ", " effective people dug ", " Clay ", TOTAL_CLAY, "YES")
Else
   If InStr(TItem, "MOULD") Then
      ' IF NO BAKERY THEN
      HEXMAPCONST.index = "FORTHKEY"
      If Not HEXMAPCONST.EOF Then
         HEXMAPCONST.MoveFirst
      End If
      HEXMAPCONST.Seek "=", TRIBESINFO![CURRENT HEX], TCLANNUMBER, "BAKERY"

      If HEXMAPCONST.NoMatch Then
         Exit Function
      End If
   End If
   Call PERFORM_COMMON("Y", "Y", "Y", 2, "ANI")
End If

ERR_PERFORM_POTTERY_CLOSE:
   Exit Function

ERR_PERFORM_POTTERY:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_POTTERY_CLOSE

End Function

Public Function PERFORM_REFINING()
On Error GoTo ERR_REFINING
TRIBE_STATUS = "PERFORM_REFINING"

' NEED TO CHECK FOR BUILDING

   BUILDING_FOUND = "N"
   
   Call CHECK_FOR_BUILDING("REFINERY")

   If BUILDING_FOUND = "N" Then
      TurnActOutPut = TurnActOutPut & "No refinery found,"
      Exit Function
   End If
   
   ' check availability of specialist.
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "REFINER")
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
          
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "REFINER", "SPECIALISTS_USED", TSpecialists)
 
   TActives = TActives + (TSpecialists * 2)
  
   Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "COKE")
  
   If RESEARCH_FOUND = "Y" Then
      'ignore
   ElseIf TItem = "COKE" Then
      GoTo ERR_REFINING_CLOSE
   End If
  
  Call PERFORM_COMMON("Y", "Y", "Y", 4, "ANI")
    
  If ModifyTable = "Y" Then
     Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "COMPANION METAL (LEAD/SILVER)")
     If RESEARCH_FOUND = "Y" Then
        If ITEM = "LEAD" Then
           Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SILVER", "ADD", CLng(NumItemsMade / 10))
        End If
     End If
   
     Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Coal Tar")
     If RESEARCH_FOUND = "Y" Then
        If ITEM = "COKE" Then
           Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Coal Tar", "ADD", CLng(NumItemsMade / 10))
        End If
     End If
  End If

ERR_REFINING_CLOSE:
   Exit Function


ERR_REFINING:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
  Resume ERR_REFINING_CLOSE
End If

End Function

Public Function PERFORM_RESEARCH()

End Function

Public Function PERFORM_SALTING()
On Error GoTo ERR_PERFORM_SALTING
TRIBE_STATUS = "PERFORM_SALTING"

'If TActives >= 10 * SALTING_LEVEL Then
'    TActives = 10 * SALTING_LEVEL
'End If

   Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
    
ERR_PERFORM_SALTING_CLOSE:
   Exit Function

ERR_PERFORM_SALTING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SALTING_CLOSE

End Function
Public Function PERFORM_SEWING()
On Error GoTo ERR_PERFORM_SEWING
TRIBE_STATUS = "PERFORM_SEWING"

   Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
    
ERR_PERFORM_SEWING_CLOSE:
   Exit Function

ERR_PERFORM_SEWING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SEWING_CLOSE

End Function

Public Function PERFORM_SIEGE_EQUIPMENT()
On Error GoTo ERR_PERFORM_SIEGE_EQUIPMENT
TRIBE_STATUS = "PERFORM_SIEGE_EQUIPMENT"

   Call PERFORM_COMMON("Y", "Y", "Y", 5, "NONE")
    
ERR_PERFORM_SIEGE_EQUIPMENT_CLOSE:
   Exit Function

ERR_PERFORM_SIEGE_EQUIPMENT:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SIEGE_EQUIPMENT_CLOSE

End Function

Function A250_PERFORM_SKILLS_CHECK()
On Error GoTo ERR_A250_PERFORM_SKILLS_CHECK
TRIBE_STATUS = "PERFORM_SKILLS_CHECK"
DebugOP "A250_PERFORM_SKILLS_CHECK"

Allskillsok = "N"
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", Skill_Tribe, TActivity

If SKILLSTABLE.NoMatch Then
   TSkillok(1) = "N"
   If TActivity = "HUNTING" Then
      TSkillok(1) = "Y"
   ElseIf TActivity = "HERDING" Then
      TSkillok(1) = "Y"
   ElseIf TActivity = "MINING" Then
      TSkillok(1) = "Y"
   ElseIf TActivity = "FURRIER" Then
      TSkillok(1) = "Y"
   End If
   
   If TJoint = "Y" Then
      TSkillok(1) = "Y"
   End If

   Skill_Level_1 = 0
   SKILL_SHORTAGE = TSkilllvl(1)
ElseIf TSkilllvl(1) <= SKILLSTABLE![SKILL LEVEL] Then
   Skill_Level_1 = SKILLSTABLE![SKILL LEVEL]
   TSkillok(1) = "Y"
ElseIf TJoint = "Y" Then
   TSkillok(1) = "Y"
   SKILL_SHORTAGE = TSkilllvl(1) - SKILLSTABLE![SKILL LEVEL]
   Skill_Level_1 = SKILLSTABLE![SKILL LEVEL]
Else
   TSkillok(1) = "N"
End If
   
If Not TSkill(2) = "FORGET" Then
   If Not IsNull(TSkill(2)) Then
      SKILLSTABLE.Seek "=", Skill_Tribe, TSkill(2)
      If SKILLSTABLE.NoMatch Then
         If TJoint = "Y" Then
            TSkillok(2) = "Y"
            Skill_Level_2 = 0
            SKILL_SHORTAGE = SKILL_SHORTAGE + TSkilllvl(2)
         Else
            TSkillok(2) = "N"
         End If
      ElseIf TSkilllvl(2) <= SKILLSTABLE![SKILL LEVEL] Then
         Skill_Level_2 = SKILLSTABLE![SKILL LEVEL]
         TSkillok(2) = "Y"
      ElseIf TJoint = "Y" Then
         TSkillok(2) = "Y"
         SKILL_SHORTAGE = SKILL_SHORTAGE + (TSkilllvl(2) - SKILLSTABLE![SKILL LEVEL])
         Skill_Level_2 = SKILLSTABLE![SKILL LEVEL]
      Else
         TSkillok(2) = "N"
      End If
    End If
End If

   If Not TSkill(3) = "FORGET" Then
      If Not IsNull(TSkill(3)) Then
         SKILLSTABLE.Seek "=", Skill_Tribe, TSkill(3)

         If SKILLSTABLE.NoMatch Then
            If TJoint = "Y" Then
               TSkillok(3) = "Y"
               SKILL_SHORTAGE = SKILL_SHORTAGE + TSkilllvl(3)
               Skill_Level_3 = 0
            Else
               TSkillok(3) = "N"
            End If
         ElseIf TSkilllvl(3) <= SKILLSTABLE![SKILL LEVEL] Then
            Skill_Level_3 = SKILLSTABLE![SKILL LEVEL]
            TSkillok(3) = "Y"
         ElseIf TJoint = "Y" Then
            TSkillok(3) = "Y"
            SKILL_SHORTAGE = SKILL_SHORTAGE + (TSkilllvl(3) - SKILLSTABLE![SKILL LEVEL])
            Skill_Level_3 = SKILLSTABLE![SKILL LEVEL]
         Else
            TSkillok(3) = "N"
         End If
       End If
   End If

   If Not TSkill(4) = "FORGET" Then
      If Not IsNull(TSkill(4)) Then
         SKILLSTABLE.Seek "=", Skill_Tribe, TSkill(4)

         If SKILLSTABLE.NoMatch Then
            If TJoint = "Y" Then
               TSkillok(4) = "Y"
               SKILL_SHORTAGE = SKILL_SHORTAGE + TSkilllvl(4)
               SKILL_LEVEL_4 = 0
            Else
               TSkillok(4) = "N"
            End If
         ElseIf TSkilllvl(4) <= SKILLSTABLE![SKILL LEVEL] Then
            SKILL_LEVEL_4 = SKILLSTABLE![SKILL LEVEL]
            TSkillok(4) = "Y"
         ElseIf TJoint = "Y" Then
            TSkillok(4) = "Y"
            SKILL_SHORTAGE = SKILL_SHORTAGE + (TSkilllvl(4) - SKILLSTABLE![SKILL LEVEL])
            SKILL_LEVEL_4 = SKILLSTABLE![SKILL LEVEL]
         Else
            TSkillok(4) = "N"
         End If
       End If
   End If

If IsNull(Skill_Level_1) Then
   MAXIMUM_ACTIVES_1 = 0
   Skill_Level_1 = 0
ElseIf Skill_Level_1 > 0 Then
   If Skill_Level_1 < 10 Then
      MAXIMUM_ACTIVES_1 = Skill_Level_1 * 10
   Else
      MAXIMUM_ACTIVES_1 = 90000
   End If
Else
   MAXIMUM_ACTIVES_1 = 0
   Skill_Level_1 = 0
End If

If IsNull(Skill_Level_2) Then
   MAXIMUM_ACTIVES_2 = 0
   Skill_Level_2 = 0
ElseIf Skill_Level_2 > 0 Then
   If Skill_Level_2 < 10 Then
      MAXIMUM_ACTIVES_2 = Skill_Level_2 * 10
   Else
      MAXIMUM_ACTIVES_2 = 90000
   End If
Else
   MAXIMUM_ACTIVES_2 = 0
   Skill_Level_2 = 0
End If

If IsNull(Skill_Level_3) Then
   MAXIMUM_ACTIVES_3 = 0
   Skill_Level_3 = 0
ElseIf Skill_Level_3 > 0 Then
   If Skill_Level_3 < 10 Then
      MAXIMUM_ACTIVES_3 = Skill_Level_3 * 10
   Else
      MAXIMUM_ACTIVES_3 = 90000
   End If
Else
   MAXIMUM_ACTIVES_3 = 0
   Skill_Level_3 = 0
End If

If IsNull(SKILL_LEVEL_4) Then
   MAXIMUM_ACTIVES_4 = 0
   SKILL_LEVEL_4 = 0
ElseIf SKILL_LEVEL_4 > 0 Then
   If SKILL_LEVEL_4 < 10 Then
      MAXIMUM_ACTIVES_4 = SKILL_LEVEL_4 * 10
   Else
      MAXIMUM_ACTIVES_4 = 90000
   End If
Else
   MAXIMUM_ACTIVES_4 = 0
   SKILL_LEVEL_4 = 0
End If

If IsNull(TSkill(2)) Then
   TSkillok(2) = "Y"
ElseIf TSkill(2) = "forget" Then
   TSkillok(2) = "Y"
End If
If IsNull(TSkill(3)) Then
   TSkillok(3) = "Y"
ElseIf TSkill(3) = "forget" Then
   TSkillok(3) = "Y"
End If
If IsNull(TSkill(4)) Then
   TSkillok(4) = "Y"
ElseIf TSkill(4) = "forget" Then
   TSkillok(4) = "Y"
End If
If TSkillok(1) = "Y" Then
   If TSkillok(2) = "Y" Then
      If TSkillok(3) = "Y" Then
         If TSkillok(4) = "Y" Then
            Allskillsok = "Y"
         Else
            Allskillsok = "N"
         End If
      Else
         Allskillsok = "N"
      End If
   Else
      Allskillsok = "N"
   End If
Else
   Allskillsok = "N"
End If

ERR_A250_PERFORM_SKILLS_CHECK_CLOSE:
   Exit Function

ERR_A250_PERFORM_SKILLS_CHECK:
   Call A999_ERROR_HANDLING
   Resume ERR_A250_PERFORM_SKILLS_CHECK_CLOSE

End Function

Public Function PERFORM_SKIN_AND_BONE()
On Error GoTo ERR_PERFORM_SKIN_AND_BONE
TRIBE_STATUS = "PERFORM_SKIN_AND_BONE"
   
' Calc Hunting Implements Used.
Call Process_Implement_Usage(TActivity, TItem, TActives, "NO")
       
   Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
   Index1 = 1
   ' TQuantity( Index1) identifys how many of each animal has been killed.
   
   If ModifyTable = "Y" Then
      If TGoods(Index1) = "CATTLE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 2))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 4))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 20))
      ElseIf TGoods(Index1) = "GOAT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 2))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 4))
      ElseIf TGoods(Index1) = "HORSE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 30))
      ElseIf TGoods(Index1) = "DOG" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 3))
      ElseIf TGoods(Index1) = "CAMEL" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 15))
      ElseIf TGoods(Index1) = "ELEPHANT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 12))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 60))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "IVORY", "ADD", (NumItemsMade * 1))
      End If
      Call UPDATE_TURNACTOUTPUT("ASNI")
   End If

ERR_PERFORM_SKIN_AND_BONE_CLOSE:
   Exit Function

ERR_PERFORM_SKIN_AND_BONE:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SKIN_AND_BONE_CLOSE

End Function

Public Function PERFORM_SKIN_AND_GUT()
On Error GoTo ERR_PERFORM_SKIN_AND_GUT
TRIBE_STATUS = "PERFORM_SKIN_AND_GUT"
' Calc Skinning & Gutting Implements Used.
Call Process_Implement_Usage(TActivity, TItem, TActives, "NO")
       
Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
Index1 = 1
   ' TQuantity( Index1) identifys how many of each animal has been killed.

If ModifyTable = "Y" Then
   If TGoods(Index1) = "CATTLE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 4))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 20))
   ElseIf TGoods(Index1) = "GOAT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 4))
   ElseIf TGoods(Index1) = "HORSE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 30))
   ElseIf TGoods(Index1) = "DOG" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 3))
   ElseIf TGoods(Index1) = "CAMEL" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 15))
   ElseIf TGoods(Index1) = "ELEPHANT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 12))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 60))
   End If
   Call UPDATE_TURNACTOUTPUT("ASNI")
End If

ERR_PERFORM_SKIN_AND_GUT_CLOSE:
   Exit Function

ERR_PERFORM_SKIN_AND_GUT:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SKIN_AND_GUT_CLOSE

End Function

Public Function PERFORM_SKIN_AND_GUT_AND_BONE()
On Error GoTo ERR_PERFORM_SKIN_AND_GUT_AND_BONE
TRIBE_STATUS = "PERFORM_SKIN_AND_GUT_AND_BONE"

' Calc Hunting Implements Used.
Call Process_Implement_Usage(TActivity, TItem, TActives, "NO")
     
ImplementsTable.index = "PRIMARYKEY"
ImplementsTable.MoveFirst
   
Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
Index1 = 1
   ' TQuantity( Index1) identifys how many of each animal has been killed.
   ' If we updated the valid animals table with skin, gut, bone, provs then we should be able to automate this and reduce the code/

If ModifyTable = "Y" Then
   If TGoods(Index1) = "CATTLE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 4))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 4))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 20))
   ElseIf TGoods(Index1) = "GOAT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 2))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 4))
   ElseIf TGoods(Index1) = "HORSE" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 30))
   ElseIf TGoods(Index1) = "DOG" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 1))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 3))
   ElseIf TGoods(Index1) = "CAMEL" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 3))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 15))
   ElseIf TGoods(Index1) = "ELEPHANT" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", NumItemsMade)
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", (NumItemsMade * 6))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GUT", "ADD", (NumItemsMade * 12))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BONES", "ADD", (NumItemsMade * 12))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", (NumItemsMade * 60))
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "IVORY", "ADD", (NumItemsMade * 1))
   End If
   Call UPDATE_TURNACTOUTPUT("ASNI")
End If

ERR_PERFORM_SKIN_AND_GUT_AND_BONE_CLOSE:
   Exit Function

ERR_PERFORM_SKIN_AND_GUT_AND_BONE:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SKIN_AND_GUT_AND_BONE_CLOSE

End Function

Public Function PERFORM_SKINNING()
On Error GoTo ERR_PERFORM_SKINNING
TRIBE_STATUS = "PERFORM_SKINNING"

   ' Calc Hunting Implements Used.
   Call Process_Implement_Usage(TActivity, TItem, TActives, "NO")

   Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
   Index1 = 1
   
   If ModifyTable = "Y" Then
      If TGoods(Index1) = "CATTLE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 1) * 2))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 20))
      ElseIf TGoods(Index1) = "GOAT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 3)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 3) * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 3) * 4))
      ElseIf TGoods(Index1) = "HORSE" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 1) * 3))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 30))
      ElseIf TGoods(Index1) = "DOG" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 1) * 1))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 3))
      ElseIf TGoods(Index1) = "CAMEL" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS * 2)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 2) * 3))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 2) * 30))
      ElseIf TGoods(Index1) = "ELEPHANT" Then
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TGoods(Index1), "SUBTRACT", TNUMOCCURS)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SKIN", "ADD", ((TNUMOCCURS * 1) * 6))
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", ((TNUMOCCURS * 1) * 60))
    End If
      Call UPDATE_TURNACTOUTPUT("ASNI")
   End If
      
ERR_PERFORM_SKINNING_CLOSE:
   Exit Function

ERR_PERFORM_SKINNING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_SKINNING_CLOSE

End Function

Public Function PERFORM_STONEWORK()
On Error GoTo ERR_PERFORM_STONEWORK
TRIBE_STATUS = "PERFORM_STONEWORK"

Dim CHECK_STONE As Long
Dim CHECK_AXE As Long
Dim CHECK_SPEAR As Long

Call PERFORM_COMMON("Y", "Y", "N", 0, "NONE")
    
Index1 = 1
   If ModifyTable = "Y" Then
      Do Until Index1 > 3
         If TGoods(Index1) = "EMPTY" Then
            Index1 = 3
         Else
            BRACKET = InStr(TGoods(Index1), "(")
            If BRACKET > 0 Then
               ITEM = Left(TGoods(Index1), (BRACKET - 1))
            Else
               ITEM = TGoods(Index1)
            End If
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", (TNUMOCCURS * TQuantity(Index1)))
         End If
         Index1 = Index1 + 1
      Loop
      
      BRACKET = InStr(TItem, "(")
      If BRACKET > 0 Then
         ITEM = Left(TItem, (BRACKET - 1))
      Else
         ITEM = TItem
      End If
      CHECK_STONE = InStr(ITEM, "STONE")
      CHECK_AXE = InStr(ITEM, "AXE")
      CHECK_SPEAR = InStr(ITEM, "SPEAR")
      
      If CHECK_STONE = 0 Then
         If Not CHECK_AXE = 0 Then
            ITEM = "STONE " & ITEM
         ElseIf Not CHECK_SPEAR = 0 Then
            ITEM = "STONE " & ITEM
         End If
      End If
      
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "ADD", NumItemsMade)
   End If

   ' update output line
   If ModifyTable = "Y" Then
      Call UPDATE_TURNACTOUTPUT("NO")
   End If

    
ERR_PERFORM_STONEWORK_CLOSE:
   Exit Function

ERR_PERFORM_STONEWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_STONEWORK_CLOSE

End Function

Public Function PERFORM_TANNING()
On Error GoTo ERR_PERFORM_TANNING
TRIBE_STATUS = "PERFORM_TANNING"

Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
    
ERR_PERFORM_TANNING_CLOSE:
   Exit Function

ERR_PERFORM_TANNING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_TANNING_CLOSE

End Function

Public Function PERFORM_WAXWORK()
On Error GoTo ERR_PERFORM_WAXWORK
TRIBE_STATUS = "PERFORM_WAXWORK"

   Call PERFORM_COMMON("Y", "Y", "Y", 4, "NONE")
    
ERR_PERFORM_WAXWORK_CLOSE:
   Exit Function

ERR_PERFORM_WAXWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_WAXWORK_CLOSE

End Function

Public Function PERFORM_WEAPONS()
On Error GoTo ERR_WEAPONS
TRIBE_STATUS = "PERFORM_WEAPONS"
   
   ' check availability of specialist.
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "WEAPONSMITH")
   
   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
          
   Call UPDATE_TRIBES_SPECIALISTS(CLAN, TRIBE, "WEAPONSMITH", "SPECIALISTS_USED", TSpecialists)
 
   TActives = TActives + (TSpecialists * 2)
   
   Call PERFORM_COMMON("Y", "Y", "Y", 4, "ANI")
    
ERR_WEAPONS_CLOSE:
   Exit Function


ERR_WEAPONS:
If (Err = 3021) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_WEAPONS_CLOSE
End If

    
End Function

Public Function PERFORM_WEAVING()
On Error GoTo ERR_PERFORM_WEAVING
TRIBE_STATUS = "PERFORM_WEAVING"

  Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "WEAVER")
   
  If TSpecialists > NO_SPECIALISTS_FOUND Then
     TSpecialists = NO_SPECIALISTS_FOUND
  End If
  
  Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "WEAVER", "SPECIALISTS_USED", TSpecialists)

  ' Actives & Specialists are combined prior to entering common function
  TActives = TActives + (TSpecialists * 2)

  Call PERFORM_COMMON("Y", "Y", "Y", 2, "ANI")
    
'if weaving Epic Tapestry then there should be a morale increase
' morale = morale + 0.04
If TItem = "EPIC TAPESTRY" Then
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   TRIBESINFO![MORALE] = TRIBESINFO![MORALE] + 0.04
   TRIBESINFO.UPDATE
End If
    
ERR_PERFORM_WEAVING_CLOSE:
   Exit Function

ERR_PERFORM_WEAVING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_WEAVING_CLOSE

End Function

Public Function PERFORM_WOODWORK()
On Error GoTo ERR_PERFORM_WOODWORK
TRIBE_STATUS = "PERFORM_WOODWORK"

  Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "WOODWORKER")
   
  If TSpecialists > NO_SPECIALISTS_FOUND Then
     TSpecialists = NO_SPECIALISTS_FOUND
  End If
  
  Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "WOODWORKER", "SPECIALISTS_USED", TSpecialists)

  ' Actives & Specialists are combined prior to entering common function
  TActives = TActives + (TSpecialists * 2)

   Call PERFORM_COMMON("Y", "Y", "Y", 3, "ANI")
    
ERR_PERFORM_WOODWORK_CLOSE:
   Exit Function

ERR_PERFORM_WOODWORK:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_WOODWORK_CLOSE

End Function

Public Function SETUP_DEFAULT()
On Error GoTo ERR_SETUP_DEFAULT
TRIBE_STATUS = "SETUP_DEFAULT"

' COPY RECORDS FROM THE COPY TO THE ACTUAL

QUERY_STRING = "Delete Process_Tribes_Activity.CLAN, Process_Tribes_Activity.TRIBE, Process_Tribes_Activity.ORDER, Process_Tribes_Activity.ACTIVITY, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity.ITEM, Process_Tribes_Activity.DISTINCTION, Process_Tribes_Activity.PEOPLE, Process_Tribes_Activity.SLAVES, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity.SPECIALISTS, Process_Tribes_Activity.JOINT, Process_Tribes_Activity.OWNING_CLAN, Process_Tribes_Activity.OWNING_TRIBE,  "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity.Number_of_Seeking_Groups, Process_Tribes_Activity.WHALE_SIZE, Process_Tribes_Activity.MINING_DIRECTION, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity.Processed "
QUERY_STRING = QUERY_STRING & " FROM Process_Tribes_Activity "
QUERY_STRING = QUERY_STRING & " WHERE (((Process_Tribes_Activity.ACTIVITY)='DEFAULT') AND ((Process_Tribes_Activity.CLAN)='" & TCLANNUMBER & "'));"
Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

QUERY_STRING = "INSERT INTO Process_Tribes_Activity ( CLAN, TRIBE, [ORDER], ACTIVITY, ITEM, DISTINCTION, PEOPLE, SLAVES, SPECIALISTS, JOINT,"
QUERY_STRING = QUERY_STRING & " OWNING_CLAN, OWNING_TRIBE, NUMBER_OF_SEEKING_GROUPS, WHALE_SIZE, MINING_DIRECTION, Processed ) "
QUERY_STRING = QUERY_STRING & " SELECT Process_Tribes_Activity_Copy.CLAN,Process_Tribes_Activity_Copy.TRIBE, Process_Tribes_Activity_Copy.ORDER,"
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity_Copy.ACTIVITY, Process_Tribes_Activity_Copy.ITEM, Process_Tribes_Activity_Copy.DISTINCTION, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity_Copy.PEOPLE, Process_Tribes_Activity_Copy.SLAVES, Process_Tribes_Activity_Copy.SPECIALISTS, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity_Copy.JOINT, Process_Tribes_Activity_Copy.OWNING_CLAN, Process_Tribes_Activity_Copy.OWNING_TRIBE, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity_Copy.NUMBER_OF_SEEKING_GROUPS, Process_Tribes_Activity_Copy.WHALE_SIZE, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Activity_Copy.MINING_DIRECTION, Process_Tribes_Activity_Copy.PROCESSED "
QUERY_STRING = QUERY_STRING & " FROM Process_Tribes_Activity_Copy  "
QUERY_STRING = QUERY_STRING & " WHERE (((Process_Tribes_Activity_Copy.CLAN)='"
QUERY_STRING = QUERY_STRING & TCLANNUMBER & "'));"
Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute


QUERY_STRING = "INSERT INTO Process_Tribes_Item_Allocation ( CLAN, TRIBE, ACTIVITY, ITEM, ITEM_USED, QUANTITY, Processed ) "
QUERY_STRING = QUERY_STRING & " SELECT Process_Tribes_Item_allocation_Copy.CLAN, Process_Tribes_Item_allocation_Copy.TRIBE, Process_Tribes_Item_allocation_Copy.ACTIVITY, "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Item_allocation_Copy.ITEM, Process_Tribes_Item_allocation_Copy.ITEM_USED, Process_Tribes_Item_allocation_Copy.QUANTITY,  "
QUERY_STRING = QUERY_STRING & " Process_Tribes_Item_allocation_Copy.PROCESSED "
QUERY_STRING = QUERY_STRING & " FROM Process_Tribes_Item_allocation_Copy "
QUERY_STRING = QUERY_STRING & " WHERE (((Process_Tribes_Item_allocation_Copy.CLAN)='"
QUERY_STRING = QUERY_STRING & TCLANNUMBER & "'));"
Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

ERR_SETUP_DEFAULT_CLOSE:
   Exit Function

ERR_SETUP_DEFAULT:
   Call A999_ERROR_HANDLING
   Resume ERR_SETUP_DEFAULT_CLOSE

End Function

Public Function UPDATE_TURNACTOUTPUT(Order As String)
On Error GoTo ERR_UPDATE_TURNACTOUTPUT
TRIBE_STATUS = "UPDATE_TURNACTOUTPUT"
      
      If TShort = "EMPTY" Then
         If Right(TurnActOutPut, 3) = "^B " Then
            TurnActOutPut = TurnActOutPut & TActives & " "
         Else
            TurnActOutPut = TurnActOutPut & ", " & TActives
         End If
      Else
         If Order = "ASNI" Then
            If TActivity = "Distilling" Then
               If Right(TurnActOutPut, 3) = "^B " Then
                  TurnActOutPut = TurnActOutPut & TActives & " effective people made " & StrConv(TShort, vbProperCase) & " " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
               Else
                  TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people made " & StrConv(TShort, vbProperCase) & " " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
               End If
            Else
               If Right(TurnActOutPut, 3) = "^B " Then
                  TurnActOutPut = TurnActOutPut & TActives & " effective people " & StrConv(TShort, vbProperCase) & " " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
               Else
                  TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people " & StrConv(TShort, vbProperCase) & " " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
               End If
            End If
         ElseIf Order = "ANI" Then
            If Right(TurnActOutPut, 3) = "^B " Then
               TurnActOutPut = TurnActOutPut & TActives & " effective people made " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
            Else
               TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people made " & NumItemsMade & " " & StrConv(TItem, vbProperCase)
            End If
         Else
            If Right(TurnActOutPut, 3) = "^B " Then
               TurnActOutPut = TurnActOutPut & StrConv(TShort, vbProperCase) & " " & NumItemsMade
            Else
               TurnActOutPut = TurnActOutPut & ", " & StrConv(TShort, vbProperCase) & " " & NumItemsMade
            End If
         End If
      End If

ERR_UPDATE_TURNACTOUTPUT_CLOSE:
   Exit Function

ERR_UPDATE_TURNACTOUTPUT:
   Call A999_ERROR_HANDLING
   Resume ERR_UPDATE_TURNACTOUTPUT_CLOSE

End Function

Public Function PERFORM_HEALING()
Attribute PERFORM_HEALING.VB_HelpID = -268435456
On Error GoTo ERR_PERFORM_HEALING
TRIBE_STATUS = "PERFORM_HEALING"

Call PERFORM_COMMON("Y", "Y", "Y", 4, "NONE")
    
ERR_PERFORM_HEALING_CLOSE:
   Exit Function

ERR_PERFORM_HEALING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_HEALING_CLOSE

End Function

Public Function Check_Turn_Output(Extra_Output1, Extra_Output2, Extra_Output3, Total_Good, Display_Actives)
On Error GoTo ERR_Check_Turn_Output
TRIBE_STATUS = "Check_Turn_Output"
   
If Total_Good = 0 Then
   If Display_Actives = "YES" Then
      TurnActOutPut = TurnActOutPut & Extra_Output1 & " " & Extra_Output2 & " " & ACTIVES_INUSE & Extra_Output3
   Else
      TurnActOutPut = TurnActOutPut & Extra_Output1 & " " & Extra_Output2 & Extra_Output3
   End If
ElseIf Display_Actives = "YES" Then
   TurnActOutPut = TurnActOutPut & Extra_Output1 & " " & ACTIVES_INUSE & " " & Extra_Output2 & " " & Total_Good & Extra_Output3
Else
   TurnActOutPut = TurnActOutPut & Extra_Output1 & " " & Extra_Output2 & " " & Total_Good & Extra_Output3
End If

ERR_Check_Turn_Output_CLOSE:
   Exit Function

ERR_Check_Turn_Output:
   Call A999_ERROR_HANDLING
   Resume ERR_Check_Turn_Output_CLOSE

End Function

Public Function Check_Available_Actives()
On Error GoTo ERR_Check_Available_Actives
TRIBE_STATUS = "Check_Available_Actives"
   
   If TActives >= (TMouths / 10) Then
      Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
      Call A150_Open_Tables("TRIBES_GENERAL_INFO")
   End If

ERR_Check_Available_Actives_CLOSE:
   Exit Function

ERR_Check_Available_Actives:
   Call A999_ERROR_HANDLING
   Resume ERR_Check_Available_Actives_CLOSE

End Function

Public Function GET_MODIFIERS()
On Error GoTo ERR_MODIFIER

TRIBE_STATUS = "GET_MODIFIERS"

If PROCESSACTIVITY![ACTIVITY] = "ENGINEERING" Then
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "STONES USED"

   If MODTABLE.NoMatch Then
      STONES_TO_USE = 5
   Else
      MODTABLE.Edit
      STONES_TO_USE = MODTABLE![AMOUNT]
   End If

ElseIf PROCESSACTIVITY![ACTIVITY] = "FORESTRY" Then
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "BARK"

   If MODTABLE.NoMatch Then
      BARK_TO_STRIP = 5
   Else
      MODTABLE.Edit
      BARK_TO_STRIP = MODTABLE![AMOUNT]
   End If

   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "LOGS"

   If MODTABLE.NoMatch Then
      LOGS_TO_CUT = 4
   Else
      MODTABLE.Edit
      LOGS_TO_CUT = MODTABLE![AMOUNT]
   End If

   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "SCRAPERS"

   If MODTABLE.NoMatch Then
      SCRAPERS_TO_USE = 0
   Else
      MODTABLE.Edit
      SCRAPERS_TO_USE = MODTABLE![AMOUNT]
   End If
ElseIf PROCESSACTIVITY![ACTIVITY] = "FURRIER" Then
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "TRAPS"

   If MODTABLE.NoMatch Then
      TRAPS_TO_USE = 5
   Else
      MODTABLE.Edit
      TRAPS_TO_USE = MODTABLE![AMOUNT]
   End If
 
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "SNARES"

   If MODTABLE.NoMatch Then
      SNARES_TO_USE = 5
   Else
      MODTABLE.Edit
      SNARES_TO_USE = MODTABLE![AMOUNT]
   End If
ElseIf PROCESSACTIVITY![ACTIVITY] = "HUNTING" Then
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "TRAPS"

   If MODTABLE.NoMatch Then
      TRAPS_TO_USE = 5
   Else
      MODTABLE.Edit
      TRAPS_TO_USE = MODTABLE![AMOUNT]
   End If
 
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "SNARES"

   If MODTABLE.NoMatch Then
      SNARES_TO_USE = 5
   Else
      MODTABLE.Edit
      SNARES_TO_USE = MODTABLE![AMOUNT]
   End If
ElseIf PROCESSACTIVITY![ACTIVITY] = "QUARRYING" Then
  
   MODTABLE.MoveFirst
   MODTABLE.Seek "=", TTRIBENUMBER, "STONES QUARRIED"

   If MODTABLE.NoMatch Then
      STONES_TO_QUARRY = 5
   Else
      MODTABLE.Edit
      STONES_TO_QUARRY = MODTABLE![AMOUNT]
   End If

End If

ERR_MODIFIER_CLOSE:
   Exit Function


ERR_MODIFIER:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_MODIFIER_CLOSE
End If

End Function


Public Function A250_GET_TERRAIN_DATA()
On Error GoTo ERR_A250_Get_Hex_Info
TRIBE_STATUS = "GET_TERRAIN_DATA"
DebugOP "A250_GET_TERRAIN_DATA"

TERRAINTABLE.MoveFirst
TERRAINTABLE.Seek "=", TRIBES_TERRAIN

If CURRENT_SEASON = "SPRING" Then
   TERRAIN_HUNTING = TERRAINTABLE![SPRING HUNTING]
ElseIf CURRENT_SEASON = "SUMMER" Then
   TERRAIN_HUNTING = TERRAINTABLE![SUMMER HUNTING]
ElseIf CURRENT_SEASON = "AUTUMN" Then
   TERRAIN_HUNTING = TERRAINTABLE![AUTUMN HUNTING]
ElseIf CURRENT_SEASON = "WINTER" Then
   TERRAIN_HUNTING = TERRAINTABLE![WINTER HUNTING]
End If

'TERRAINTABLE.MoveFirst
'TERRAINTABLE.Seek "=", GOODS_TRIBES_TERRAIN

If CURRENT_SEASON = "SPRING" Then
   TERRAIN_HERDING_GROUP_1 = TERRAINTABLE![SPRING HERDING GROUP 1]
   TERRAIN_HERDING_GROUP_2 = TERRAINTABLE![SPRING HERDING GROUP 2]
   TERRAIN_HERDING_GROUP_3 = TERRAINTABLE![SPRING HERDING GROUP 3]
ElseIf CURRENT_SEASON = "SUMMER" Then
   TERRAIN_HERDING_GROUP_1 = TERRAINTABLE![SUMMER HERDING GROUP 1]
   TERRAIN_HERDING_GROUP_2 = TERRAINTABLE![SUMMER HERDING GROUP 2]
   TERRAIN_HERDING_GROUP_3 = TERRAINTABLE![SUMMER HERDING GROUP 3]
ElseIf CURRENT_SEASON = "AUTUMN" Then
   TERRAIN_HERDING_GROUP_1 = TERRAINTABLE![AUTUMN HERDING GROUP 1]
   TERRAIN_HERDING_GROUP_2 = TERRAINTABLE![AUTUMN HERDING GROUP 2]
   TERRAIN_HERDING_GROUP_3 = TERRAINTABLE![AUTUMN HERDING GROUP 3]
ElseIf CURRENT_SEASON = "WINTER" Then
   TERRAIN_HERDING_GROUP_1 = TERRAINTABLE![WINTER HERDING GROUP 1]
   TERRAIN_HERDING_GROUP_2 = TERRAINTABLE![WINTER HERDING GROUP 2]
   TERRAIN_HERDING_GROUP_3 = TERRAINTABLE![WINTER HERDING GROUP 3]
End If

ERR_A250_Get_Hex_Info_CLOSE:
   Exit Function


ERR_A250_Get_Hex_Info:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Hex_Info_CLOSE
End If
End Function

Public Function A250_GET_SEASON_DATA()
On Error GoTo ERR_A250_Get_Hex_Info

TRIBE_STATUS = "GET_SEASON_DATA"
DebugOP "A250_GET_SEASON_DATA"

CURRENT_SEASON = GET_SEASON(Current_Turn)

SEASONTABLE.MoveFirst
SEASONTABLE.Seek "=", CURRENT_SEASON

SEASON_HONEY = SEASONTABLE![HONEY]
SEASON_WAX = SEASONTABLE![WAX]
COASTAL_FISHING = SEASONTABLE![FISHING - COASTAL]
OCEAN_FISHING = SEASONTABLE![FISHING - OCEAN]
FURRIER_SKINS = SEASONTABLE![FURRIER - SKINS]
FURRIER_FURS = SEASONTABLE![FURRIER - FURS]

ERR_A250_Get_Hex_Info_CLOSE:
   Exit Function


ERR_A250_Get_Hex_Info:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Hex_Info_CLOSE
End If
End Function


Public Function A250_GET_WIND_DATA()
On Error GoTo ERR_A250_Get_Wind_Data
TRIBE_STATUS = "GET_WIND_DATA"
DebugOP "A250_GET_WIND_DATA"

WINDTABLE.MoveFirst
WINDTABLE.Seek "=", CURRENT_WIND

WIND_FISHING = WINDTABLE![FISHING]

ERR_A250_Get_Wind_Data_CLOSE:
   Exit Function

ERR_A250_Get_Wind_Data:
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Wind_Data_CLOSE

End Function

Public Function A250_Get_Tribes_Skills()
On Error GoTo ERR_A250_Get_Tribes_Skills
TRIBE_STATUS = "GET_SKILLS_DATA"
DebugOP "A250_Get_Tribes_Skills"

SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", Skill_Tribe, "APIARISM"

If SKILLSTABLE.NoMatch Then
   APIARISM_LEVEL = 0
Else
   APIARISM_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "BONING"

If SKILLSTABLE.NoMatch Then
   BONING_LEVEL = 0
Else
   BONING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "COMBAT"

If SKILLSTABLE.NoMatch Then
   COMBAT_LEVEL = 0
Else
   COMBAT_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "DIPLOMACY"

If SKILLSTABLE.NoMatch Then
   DIPLOMACY_LEVEL = 0
Else
   DIPLOMACY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "FARMING"

If SKILLSTABLE.NoMatch Then
   FARMING_LEVEL = 0
Else
   FARMING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "FISHING"

If SKILLSTABLE.NoMatch Then
   FISHING_LEVEL = 0
Else
   FISHING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "FLENSING"

If SKILLSTABLE.NoMatch Then
   FLENSING_LEVEL = 0
Else
   FLENSING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "FORESTRY"

If SKILLSTABLE.NoMatch Then
   FORESTRY_LEVEL = 0
Else
   FORESTRY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "FURRIER"

If SKILLSTABLE.NoMatch Then
   FURRIER_LEVEL = 0
Else
   FURRIER_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "GUTTING"

If SKILLSTABLE.NoMatch Then
   GUTTING_LEVEL = 0
Else
   GUTTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "HEALING"

If SKILLSTABLE.NoMatch Then
   HEALING_LEVEL = 0
Else
   HEALING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "HERDING"

If SKILLSTABLE.NoMatch Then
   HERDING_LEVEL = 0
Else
   HERDING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "HUNTING"

If SKILLSTABLE.NoMatch Then
   HUNTING_LEVEL = 0
Else
   HUNTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "MINING"

If SKILLSTABLE.NoMatch Then
   MINING_LEVEL = 0
Else
   MINING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "PEELING"

If SKILLSTABLE.NoMatch Then
   PEELING_LEVEL = 0
Else
   PEELING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "POLITICS"

If SKILLSTABLE.NoMatch Then
   POLITICS_LEVEL = 0
Else
   POLITICS_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "RELIGION"

If SKILLSTABLE.NoMatch Then
   RELIGION_LEVEL = 0
Else
   RELIGION_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SALTING"

If SKILLSTABLE.NoMatch Then
   SALTING_LEVEL = 0
Else
   SALTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SANITATION"

If SKILLSTABLE.NoMatch Then
   SANITATION_LEVEL = 0
Else
   SANITATION_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SCOUTING"

If SKILLSTABLE.NoMatch Then
   SCOUTING_LEVEL = 0
Else
   SCOUTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SEEKING"

If SKILLSTABLE.NoMatch Then
   SEEKING_LEVEL = 0
Else
   SEEKING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SKINNING"

If SKILLSTABLE.NoMatch Then
   SKINNING_LEVEL = 0
Else
   SKINNING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "SLAVERY"

If SKILLSTABLE.NoMatch Then
   SLAVERY_LEVEL = 0
Else
   SLAVERY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Seek "=", Skill_Tribe, "WHALING"

If SKILLSTABLE.NoMatch Then
   WHALING_LEVEL = 0
Else
   WHALING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

ERR_A250_Get_Tribes_Skills_CLOSE:
   Exit Function


ERR_A250_Get_Tribes_Skills:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Tribes_Skills_CLOSE
End If
End Function

Public Function SET_SKILL_LEVEL_1(Skill_Level As Long)
On Error GoTo ERR_SET_SKILL_LEVEL_1
TRIBE_STATUS = "SET_SKILL_LEVEL_1"

If Skill_Level = 0 Then
   Skill_Level_1 = 0
   MAXIMUM_ACTIVES_1 = 0
Else
   Skill_Level_1 = Skill_Level
   If Skill_Level_1 >= 10 Then
      MAXIMUM_ACTIVES_1 = 80000
   Else
      MAXIMUM_ACTIVES_1 = Skill_Level_1 * 10
   End If
End If

ERR_SET_SKILL_LEVEL_1_CLOSE:
   Exit Function

ERR_SET_SKILL_LEVEL_1:
   Call A999_ERROR_HANDLING
   Resume ERR_SET_SKILL_LEVEL_1_CLOSE

End Function

Public Function SET_SKILL_LEVEL_2(Skill_Level As Long)
On Error GoTo ERR_SET_SKILL_LEVEL_2
TRIBE_STATUS = "SET_SKILL_LEVEL_2"

If Skill_Level = 0 Then
   Skill_Level_2 = 0
   MAXIMUM_ACTIVES_2 = 0
Else
   Skill_Level_2 = Skill_Level
   If Skill_Level_2 >= 10 Then
      MAXIMUM_ACTIVES_2 = 80000
   Else
      MAXIMUM_ACTIVES_2 = Skill_Level_2 * 10
   End If
End If

ERR_SET_SKILL_LEVEL_2_CLOSE:
   Exit Function

ERR_SET_SKILL_LEVEL_2:
   Call A999_ERROR_HANDLING
   Resume ERR_SET_SKILL_LEVEL_2_CLOSE

End Function

Public Function SET_SKILL_LEVEL_3(Skill_Level As Long)
On Error GoTo ERR_SET_SKILL_LEVEL_3
TRIBE_STATUS = "SET_SKILL_LEVEL_3"
  
If Skill_Level = 0 Then
   Skill_Level_3 = 0
   MAXIMUM_ACTIVES_3 = 0
Else
   Skill_Level_3 = Skill_Level
   If Skill_Level_3 >= 10 Then
      MAXIMUM_ACTIVES_3 = 80000
   Else
      MAXIMUM_ACTIVES_3 = Skill_Level_3 * 10
   End If
End If

ERR_SET_SKILL_LEVEL_3_CLOSE:
   Exit Function

ERR_SET_SKILL_LEVEL_3:
   Call A999_ERROR_HANDLING
   Resume ERR_SET_SKILL_LEVEL_3_CLOSE

End Function

Public Function Process_Animal_Eating()
On Error GoTo ERR_Process_Animal_Eating
TRIBE_STATUS = "Process_Animal_Eating"

TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "WARHORSES"

If TRIBESGOODS.NoMatch Then
   WARHORSES = 0
Else
   WARHORSES = TRIBESGOODS![ITEM_NUMBER]
End If

   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "GRAIN"
   GRAINREQ = 0
   FODDERREQ = 0
   If WARHORSES > 0 Then
      GRAINREQ = WARHORSES * 8
      FODDERREQ = WARHORSES * 10
   End If
  
   '
   '
   '
   ' UPDATE OUTPUT WITH FACT OF WARHORSES FED
   ' GET LAST DETAIL LINE
   '
   '
   '
   '
   Do While GRAINREQ > 0
      If TRIBESGOODS.NoMatch Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "FODDER"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "WARHORSES"
            If Not TRIBESGOODS.NoMatch Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (GRAINREQ / 8)
               TRIBESGOODS.UPDATE
            End If
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "HORSE"
            If Not TRIBESGOODS.NoMatch Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (GRAINREQ / 8)
               TRIBESGOODS.UPDATE
            End If
            GRAINREQ = 0
            FODDERREQ = 0
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 10
            TRIBESGOODS.UPDATE
         End If
      ElseIf TRIBESGOODS![ITEM] = "GRAIN" Then
         If TRIBESGOODS![ITEM_NUMBER] > 0 Then
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 8
            TRIBESGOODS.UPDATE
         Else
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "FODDER"
            If TRIBESGOODS.NoMatch Then
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "WARHORSES"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (GRAINREQ / 8)
               TRIBESGOODS.UPDATE
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "HORSE"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (GRAINREQ / 8)
               TRIBESGOODS.UPDATE
               GRAINREQ = 0
               FODDERREQ = 0
            Else
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 8
               TRIBESGOODS.UPDATE
            End If
         End If
      Else
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "FODDER"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "WARHORSES"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (GRAINREQ / 8)
            TRIBESGOODS.UPDATE
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "HORSE"
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (GRAINREQ / 8)
            TRIBESGOODS.UPDATE
            GRAINREQ = 0
            FODDERREQ = 0
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 10
            TRIBESGOODS.UPDATE
         End If
      End If
      GRAINREQ = GRAINREQ - 8
      FODDERREQ = FODDERREQ - 10
   Loop
 


ERR_Process_Animal_Eating_CLOSE:
   Exit Function

ERR_Process_Animal_Eating:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Animal_Eating_CLOSE

End Function

Public Function Process_Other_Final_Activities()
On Error GoTo ERR_Process_Other_Final_Activities
TRIBE_STATUS = "Process_Other_Final_Activities"

Dim SEQ_NUMBER As Long
Dim TRIBE As String
Dim CONSTRUCTION As String
Dim LOGS As Long
Dim STONES As Long
Dim COAL As Long
Dim BRASS As Long
Dim BRONZE As Long
Dim COPPER As Long
Dim IRON As Long
Dim LEAD As Long
Dim CLOTH As Long
Dim LEATHER As Long
Dim ROPES As Long
Dim LOGS_H As Long
Dim QUERY As String
Dim ConstructionTable2 As Recordset

'Convert bricks into stones

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "BRICKS"
If Not TRIBESGOODS.NoMatch Then
   BRICKS_CREATED = TRIBESGOODS![ITEM_NUMBER]
   TRIBESGOODS.Delete
   TRIBESGOODS.UPDATE
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "STONE"
   If TRIBESGOODS.NoMatch Then
      TRIBESGOODS.AddNew
      TRIBESGOODS![CLAN] = TCLANNUMBER
      TRIBESGOODS![TRIBE] = GOODS_TRIBE
      TRIBESGOODS![ITEM_TYPE] = "RAW"
      TRIBESGOODS![ITEM] = "STONE"
      TRIBESGOODS![ITEM_NUMBER] = BRICKS_CREATED
      TRIBESGOODS.UPDATE
   Else
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + BRICKS_CREATED
      TRIBESGOODS.UPDATE
   End If
End If

'Perform a morale throw

If TRIBESINFO![MORALE] < 1 Then
   If MORALELOSS = "N" Then
      If Len(TTRIBENUMBER) > 4 Then
         DICE_TRIBE = Left(TTRIBENUMBER, 4)
      ElseIf Left(TTRIBENUMBER, 1) = "B" Then
         DICE_TRIBE = CLng(TCLANNUMBER)
      ElseIf Left(TTRIBENUMBER, 1) = "M" Then
         DICE_TRIBE = CLng(TCLANNUMBER)
      Else
         DICE_TRIBE = CLng(TTRIBENUMBER)
      End If
      
      DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 1, 0)
      
      If DICE1 >= 50 Then
         Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER)
         Call A150_Open_Tables("TRIBES_GENERAL_INFO")
      End If
   End If
End If

'Update Months Trading Post open figures

If Not HEXMAPCONST.EOF Then
   HEXMAPCONST.index = "PRIMARYKEY"
   HEXMAPCONST.MoveFirst
End If
HEXMAPCONST.Seek "=", TRIBESINFO![CURRENT HEX], TCLANNUMBER, TTRIBENUMBER, "MEETING HOUSE"
VILLAGE_FOUND = "N"
  
If HEXMAPCONST.NoMatch Then
   VILLAGE_FOUND = "N"
   HEXMAPCONST.MoveFirst
ElseIf HEXMAPCONST![CLAN] = TCLANNUMBER Then
   VILLAGE_FOUND = "Y"
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", TRIBESINFO![CURRENT HEX], TCLANNUMBER, TTRIBENUMBER, "TRADING POST"
   If Not HEXMAPCONST.NoMatch Then
      If HEXMAPCONST![1] > 0 Then
         HEXMAPCONST.MoveFirst
         HEXMAPCONST.Seek "=", TRIBESINFO![CURRENT HEX], TCLANNUMBER, TTRIBENUMBER, "Months TP Open"
         If HEXMAPCONST.NoMatch Then
            HEXMAPCONST.AddNew
            HEXMAPCONST![MAP] = TRIBESINFO![CURRENT HEX]
            HEXMAPCONST![CLAN] = TCLANNUMBER
            HEXMAPCONST![TRIBE] = TTRIBENUMBER
            HEXMAPCONST![CONSTRUCTION] = "MONTHS TP OPEN"
            HEXMAPCONST![1] = 1
            HEXMAPCONST.UPDATE
         Else
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
         End If
      End If
   End If
   HEXMAPCONST.MoveFirst
Else
   VILLAGE_FOUND = "N"
   HEXMAPCONST.MoveFirst
End If

' update for training specialists
TribesSpecialists.index = "PRIMARYKEY"
TribesSpecialists.MoveFirst

If Not TribesSpecialists.EOF Then
   TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, "TRAINING"
    If Not TribesSpecialists.NoMatch Then
            TribesSpecialists.Edit
            TribesSpecialists![NUMBER_OF_TURNS_TRAINING] = TribesSpecialists![NUMBER_OF_TURNS_TRAINING] + 1
            TribesSpecialists.UPDATE
    End If
End If

' update for breeding new queens
' restrict to sping & summer only

If Left(Current_Turn, 2) = 1 Then

    Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Breed new Queens")
    If RESEARCH_FOUND = "Y" Then
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "HIVE"
      If TRIBESGOODS.NoMatch Then
         ' DO NOTHING
      Else
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + 24
         TRIBESGOODS.UPDATE
      End If
   End If
End If

' update under construction to fix sequencing problems.

ConstructionTable.MoveFirst

Set ConstructionTable2 = TVDB.OpenRecordset("Under_Construction_TEMP")
ConstructionTable2.index = "PRIMARYKEY"

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM UNDER_CONSTRUCTION_TEMP;")
qdfCurrent.Execute

Do
   If ConstructionTable.EOF Then
      Exit Do
   End If
   
   TRIBE = ConstructionTable![TRIBE]
   CONSTRUCTION = ConstructionTable![CONSTRUCTION]
   LOGS = ConstructionTable![LOGS]
   STONES = ConstructionTable![STONES]
   COAL = ConstructionTable![COAL]
   BRASS = ConstructionTable![BRASS]
   BRONZE = ConstructionTable![BRONZE]
   COPPER = ConstructionTable![COPPER]
   IRON = ConstructionTable![IRON]
   LEAD = ConstructionTable![LEAD]
   CLOTH = ConstructionTable![CLOTH]
   LEATHER = ConstructionTable![LEATHER]
   ROPES = ConstructionTable![ROPES]
   LOGS_H = ConstructionTable![LOG/H]
   ConstructionTable.Delete
      
'  ConstructionTable2.MoveFirst
   ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 1
   If ConstructionTable2.NoMatch Then
      SEQ_NUMBER = 1
   Else
      ConstructionTable2.MoveFirst
      ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 2
      If ConstructionTable2.NoMatch Then
         SEQ_NUMBER = 2
      Else
         ConstructionTable2.MoveFirst
         ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 3
         If ConstructionTable2.NoMatch Then
            SEQ_NUMBER = 3
         Else
            ConstructionTable2.MoveFirst
            ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 4
            If ConstructionTable2.NoMatch Then
               SEQ_NUMBER = 4
            Else
               ConstructionTable2.MoveFirst
               ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 5
               If ConstructionTable2.NoMatch Then
                  SEQ_NUMBER = 5
               Else
                  ConstructionTable2.MoveFirst
                  ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 6
                  If ConstructionTable2.NoMatch Then
                     SEQ_NUMBER = 6
                  Else
                     ConstructionTable2.MoveFirst
                     ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 7
                     If ConstructionTable2.NoMatch Then
                        SEQ_NUMBER = 7
                     Else
                        ConstructionTable2.MoveFirst
                        ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 8
                        If ConstructionTable2.NoMatch Then
                           SEQ_NUMBER = 8
                        Else
                           ConstructionTable2.MoveFirst
                           ConstructionTable2.Seek "=", TRIBE, CONSTRUCTION, 9
                           If ConstructionTable2.NoMatch Then
                              SEQ_NUMBER = 9
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
          
   ConstructionTable2.AddNew
   ConstructionTable2![TRIBE] = TRIBE
   ConstructionTable2![CONSTRUCTION] = CONSTRUCTION
   ConstructionTable2![SEQ NUMBER] = SEQ_NUMBER
   ConstructionTable2![LOGS] = LOGS
   ConstructionTable2![STONES] = STONES
   ConstructionTable2![COAL] = COAL
   ConstructionTable2![BRASS] = BRASS
   ConstructionTable2![BRONZE] = BRONZE
   ConstructionTable2![COPPER] = COPPER
   ConstructionTable2![IRON] = IRON
   ConstructionTable2![LEAD] = LEAD
   ConstructionTable2![CLOTH] = CLOTH
   ConstructionTable2![LEATHER] = LEATHER
   ConstructionTable2![ROPES] = ROPES
   ConstructionTable2![LOG/H] = LOGS_H
   ConstructionTable2.UPDATE
   ConstructionTable.MoveFirst

Loop

ConstructionTable2.MoveFirst
Do
   If ConstructionTable2.EOF Then
      Exit Do
   End If
   
   TRIBE = ConstructionTable2![TRIBE]
   CONSTRUCTION = ConstructionTable2![CONSTRUCTION]
   LOGS = ConstructionTable2![LOGS]
   STONES = ConstructionTable2![STONES]
   COAL = ConstructionTable2![COAL]
   BRASS = ConstructionTable2![BRASS]
   BRONZE = ConstructionTable2![BRONZE]
   COPPER = ConstructionTable2![COPPER]
   IRON = ConstructionTable2![IRON]
   LEAD = ConstructionTable2![LEAD]
   CLOTH = ConstructionTable2![CLOTH]
   LEATHER = ConstructionTable2![LEATHER]
   ROPES = ConstructionTable2![ROPES]
   LOGS_H = ConstructionTable2![LOG/H]
   ConstructionTable2.Delete
      
   ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 1
   If ConstructionTable.NoMatch Then
      SEQ_NUMBER = 1
   Else
      ConstructionTable.MoveFirst
      ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 2
      If ConstructionTable.NoMatch Then
         SEQ_NUMBER = 2
      Else
         ConstructionTable.MoveFirst
         ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 3
         If ConstructionTable.NoMatch Then
            SEQ_NUMBER = 3
         Else
            ConstructionTable.MoveFirst
            ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 4
            If ConstructionTable.NoMatch Then
               SEQ_NUMBER = 4
            Else
               ConstructionTable.MoveFirst
               ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 5
               If ConstructionTable.NoMatch Then
                  SEQ_NUMBER = 5
               Else
                  ConstructionTable.MoveFirst
                  ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 6
                  If ConstructionTable.NoMatch Then
                     SEQ_NUMBER = 6
                  Else
                     ConstructionTable.MoveFirst
                     ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 7
                     If ConstructionTable.NoMatch Then
                        SEQ_NUMBER = 7
                     Else
                       ConstructionTable.MoveFirst
                       ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 8
                       If ConstructionTable.NoMatch Then
                          SEQ_NUMBER = 8
                       Else
                          ConstructionTable.MoveFirst
                          ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 9
                          If ConstructionTable.NoMatch Then
                             SEQ_NUMBER = 9
                          End If
                       End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
          
   ConstructionTable.AddNew
   ConstructionTable![TRIBE] = TRIBE
   ConstructionTable![CONSTRUCTION] = CONSTRUCTION
   ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
   ConstructionTable![LOGS] = LOGS
   ConstructionTable![STONES] = STONES
   ConstructionTable![COAL] = COAL
   ConstructionTable![BRASS] = BRASS
   ConstructionTable![BRONZE] = BRONZE
   ConstructionTable![COPPER] = COPPER
   ConstructionTable![IRON] = IRON
   ConstructionTable![LEAD] = LEAD
   ConstructionTable![CLOTH] = CLOTH
   ConstructionTable![LEATHER] = LEATHER
   ConstructionTable![ROPES] = ROPES
   ConstructionTable![LOG/H] = LOGS_H
   ConstructionTable.UPDATE
   ConstructionTable2.MoveFirst
   
Loop

ConstructionTable2.Close

ERR_Process_Other_Final_Activities_CLOSE:
   Exit Function

ERR_Process_Other_Final_Activities:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Other_Final_Activities_CLOSE

End Function

Public Function Populate_Implement_Usage_Table()
' This is used in the Clean_Up_and_Reset process.
' This occurs at the start of processing.
On Error GoTo ERR_POPULATE
TRIBE_STATUS = "Populate_Implement_Usage_Table"

Set TVWKSPACE = DBEngine.Workspaces(0)

Call A150_Open_Tables("all")

TRIBESGOODS.index = "TERTIARYKEY"
TRIBESGOODS.MoveFirst

Do Until TRIBESINFO.EOF
   TCLANNUMBER = TRIBESINFO![CLAN]
   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      GOODS_TRIBE = TRIBESINFO![TRIBE]
   End If
   
   ImplementsTable.MoveFirst
   ImplementUsage.MoveFirst
   ImplementUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, ImplementsTable![IMPLEMENT]
   If ImplementUsage.NoMatch Then
      Do Until ImplementsTable.EOF
         ' READ THE GOODS TABLE
      
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, ImplementsTable![IMPLEMENT]
   
         If Not TRIBESGOODS.NoMatch Then
            ImplementUsage.AddNew
            ImplementUsage![CLAN] = TCLANNUMBER
            ImplementUsage![TRIBE] = GOODS_TRIBE
            ImplementUsage![IMPLEMENT] = ImplementsTable![IMPLEMENT]
            ImplementUsage![total_available] = TRIBESGOODS![ITEM_NUMBER]
            ImplementUsage![Number_Used] = 0
            ImplementUsage.UPDATE
         End If
         
         ImplementsTable.MoveNext
     
      Loop
   End If
   TRIBESINFO.MoveNext
Loop

ERR_POP_CLOSE:
   Exit Function

ERR_POPULATE:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_POP_CLOSE
End If

End Function
Public Function Process_Population_Growth()
On Error GoTo ERR_Process_Population_Growth
TRIBE_STATUS = "Process_Population_Growth"

' Calculate the population increase
' get smallest of wp, ap, inap
    
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
TRIBESINFO.Edit
    
AvailableSerfs = 0
    
HEXMAPPOLITICS.MoveFirst

Do
     If HEXMAPPOLITICS![PL_CLAN] = TCLANNUMBER And HEXMAPPOLITICS![PL_TRIBE] = TTRIBENUMBER Then
         If IsNull(HEXMAPPOLITICS![POPULATION]) Then
             If HEXMAPPOLITICS![POP_INCREASED] = "N" Then
                  HEXMAPPOLITICS.Edit
                  HEXMAPPOLITICS![POPULATION] = 400
                  HEXMAPPOLITICS![POP_INCREASED] = "Y"
                  HEXMAPPOLITICS.UPDATE
             End If
         Else
         If HEXMAPPOLITICS![POP_INCREASED] = "N" Or IsNull(HEXMAPPOLITICS![POP_INCREASED]) Then
               If HEXMAPPOLITICS![POPULATION] > 0 Then
                  HEX_POP_GROWTH = (0.01 * HEALING_LEVEL)
                  HEX_POP_GROWTH = HEX_POP_GROWTH + 1
         
                  HEXMAPPOLITICS.Edit
                  AvailableSerfs = AvailableSerfs + CLng(HEXMAPPOLITICS![POPULATION] * HEX_POP_GROWTH)
                  HEXMAPPOLITICS![POPULATION] = CLng(HEXMAPPOLITICS![POPULATION] + HEX_POP_GROWTH)
                  HEXMAPPOLITICS![POP_INCREASED] = "Y"
                  HEXMAPPOLITICS.UPDATE
               Else
                  HEXMAPPOLITICS.Edit
                  HEXMAPPOLITICS![POPULATION] = 400
                  AvailableSerfs = AvailableSerfs + CLng(HEXMAPPOLITICS![POPULATION] * HEX_POP_GROWTH)
                  HEXMAPPOLITICS![POP_INCREASED] = "Y"
                  HEXMAPPOLITICS.UPDATE
               End If
            End If
         End If
     End If
     HEXMAPPOLITICS.MoveNext
     If HEXMAPPOLITICS.EOF Then
        Exit Do
     End If
Loop
   
If IsNull(TRIBESINFO![POP TRIBE]) Then
   TRIBES_POP_TRIBE = TTRIBENUMBER
Else
   TRIBES_POP_TRIBE = TRIBESINFO![POP TRIBE]
End If
    
' perform calc of new monthly figure
    
Set PopTable = TVDBGM.OpenRecordset("Population_Increase")
PopTable.index = "PRIMARYKEY"
PopTable.MoveFirst
PopTable.Seek "=", TCLANNUMBER, TTRIBENUMBER
If PopTable.NoMatch Then
   PopTable.AddNew
   PopTable![CLAN] = TCLANNUMBER
   PopTable![TRIBE] = TTRIBENUMBER
   PopTable.UPDATE
   PopTable.Seek "=", TCLANNUMBER, TTRIBENUMBER
End If

PopTable.Edit
TOldWarriors = TRIBESINFO!WARRIORS
If TOldWarriors < 0 Then
   TOldWarriors = 0
End If
TOldActives = TRIBESINFO!ACTIVES
If TOldActives < 0 Then
   TOldActives = 0
End If
TOldInactives = TRIBESINFO!INACTIVES
If TOldInactives < 0 Then
   TOldInactives = 0
End If
If TOldWarriors < TOldActives Then
   If TOldWarriors < TOldInactives Then
      TOldPopFigure = TOldWarriors
   Else
      TOldPopFigure = TOldInactives
   End If
ElseIf TOldActives < TOldInactives Then
   TOldPopFigure = TOldActives
Else
   TOldPopFigure = TOldInactives
End If
    
TribesSpecialists.index = "SECONDARYKEY"
TribesSpecialists.MoveFirst
TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER
    
If TribesSpecialists.NoMatch Then
   ' DO NOTHING
Else
   Do Until Not (TribesSpecialists![TRIBE] = TTRIBENUMBER)
      TOldPopFigure = TOldPopFigure + TribesSpecialists![SPECIALISTS]
      TribesSpecialists.MoveNext
      If TribesSpecialists.EOF Then
         Exit Do
      End If
   Loop
End If
    
' get building info

Call CHECK_FOR_BUILDING("HOSPITAL")
Call CHECK_FOR_BUILDING("SEWER")

'GET RESEARCH INFO
RESEARCH_FOUND = "N"
Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "MIDWIFERY")
If RESEARCH_FOUND = "Y" Then
   MIDWIFERY_FOUND = "Y"
Else
   MIDWIFERY_FOUND = "N"
End If

Tribe_Population_Increase = MORALE / 150

If HOSPITAL_FOUND > 0 Then
   Tribe_Population_Increase = Tribe_Population_Increase + (0.0005 * HEALING_LEVEL)
End If

If SEWER_FOUND > 0 Then
   Tribe_Population_Increase = Tribe_Population_Increase + (0.0005 * SANITATION_LEVEL)
End If

If MIDWIFERY_FOUND = "Y " Then
   Tribe_Population_Increase = Tribe_Population_Increase + (0.001 * HEALING_LEVEL)
End If

Call A150_Open_Tables("MODIFIERS")
MODTABLE.MoveFirst
MODTABLE.Seek "=", TTRIBENUMBER, "POPULATION INCREASE"
   
If MODTABLE.NoMatch Then
   ' DO NOTHING
Else
   MODTABLE.Edit
   Tribe_Population_Increase = Tribe_Population_Increase + MODTABLE![AMOUNT]
End If
   
AvailableSerfs = CLng(AvailableSerfs * 0.003)

If TURN_NUMBER = "TURN01" Then
   Current_Increase = PopTable![TURN01] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN02" Then
   Current_Increase = PopTable![TURN02] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN03" Then
   Current_Increase = PopTable![TURN03] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN04" Then
   Current_Increase = PopTable![TURN04] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN05" Then
   Current_Increase = PopTable![TURN05] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN06" Then
   Current_Increase = PopTable![TURN06] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN07" Then
   Current_Increase = PopTable![TURN07] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN08" Then
   Current_Increase = PopTable![TURN08] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN09" Then
   Current_Increase = PopTable![TURN09] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN10" Then
   Current_Increase = PopTable![TURN10] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN11" Then
   Current_Increase = PopTable![TURN11] + AvailableSerfs
ElseIf TURN_NUMBER = "TURN12" Then
   Current_Increase = PopTable![TURN12] + AvailableSerfs
End If
    
If TURN_NUMBER = "TURN01" Then
   PopTable![TURN10] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN10] = CLng(PopTable![TURN10] * PopTable![RELIGION])
   PopTable![TURN10] = CLng(PopTable![TURN10] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN10]
   End If
ElseIf TURN_NUMBER = "TURN02" Then
   PopTable![TURN11] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN11] = CLng(PopTable![TURN11] * PopTable![RELIGION])
   PopTable![TURN11] = CLng(PopTable![TURN11] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN11]
   End If
ElseIf TURN_NUMBER = "TURN03" Then
   PopTable![TURN12] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN12] = CLng(PopTable![TURN12] * PopTable![RELIGION])
   PopTable![TURN12] = CLng(PopTable![TURN12] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN12]
   End If
ElseIf TURN_NUMBER = "TURN04" Then
   PopTable![TURN01] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN01] = CLng(PopTable![TURN01] * PopTable![RELIGION])
   PopTable![TURN01] = CLng(PopTable![TURN01] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN01]
   End If
ElseIf TURN_NUMBER = "TURN05" Then
   PopTable![TURN02] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN02] = CLng(PopTable![TURN02] * PopTable![RELIGION])
   PopTable![TURN02] = CLng(PopTable![TURN02] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN02]
   End If
ElseIf TURN_NUMBER = "TURN06" Then
   PopTable![TURN03] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN03] = CLng(PopTable![TURN03] * PopTable![RELIGION])
   PopTable![TURN03] = CLng(PopTable![TURN03] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN03]
   End If
ElseIf TURN_NUMBER = "TURN07" Then
   PopTable![TURN04] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
   PopTable![TURN04] = CLng(PopTable![TURN04] * PopTable![RELIGION])
   PopTable![TURN04] = CLng(PopTable![TURN04] * PopTable![HOSPITAL])
   If Current_Increase = 0 Then
      Current_Increase = PopTable![TURN04]
   End If
ElseIf TURN_NUMBER = "TURN08" Then
       PopTable![TURN05] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
       PopTable![TURN05] = CLng(PopTable![TURN05] * PopTable![RELIGION])
       PopTable![TURN05] = CLng(PopTable![TURN05] * PopTable![HOSPITAL])
       If Current_Increase = 0 Then
          Current_Increase = PopTable![TURN05]
       End If
    ElseIf TURN_NUMBER = "TURN09" Then
       PopTable![TURN06] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
       PopTable![TURN06] = CLng(PopTable![TURN06] * PopTable![RELIGION])
       PopTable![TURN06] = CLng(PopTable![TURN06] * PopTable![HOSPITAL])
       If Current_Increase = 0 Then
          Current_Increase = PopTable![TURN06]
       End If
    ElseIf TURN_NUMBER = "TURN10" Then
       PopTable![TURN07] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
       PopTable![TURN07] = CLng(PopTable![TURN07] * PopTable![RELIGION])
       PopTable![TURN07] = CLng(PopTable![TURN07] * PopTable![HOSPITAL])
       If Current_Increase = 0 Then
          Current_Increase = PopTable![TURN07]
       End If
    ElseIf TURN_NUMBER = "TURN11" Then
       PopTable![TURN08] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
       PopTable![TURN08] = CLng(PopTable![TURN08] * PopTable![RELIGION])
       PopTable![TURN08] = CLng(PopTable![TURN08] * PopTable![HOSPITAL])
       If Current_Increase = 0 Then
          Current_Increase = PopTable![TURN08]
       End If
    ElseIf TURN_NUMBER = "TURN12" Then
       PopTable![TURN09] = CLng((TOldPopFigure * 3) * Tribe_Population_Increase)
       PopTable![TURN09] = CLng(PopTable![TURN09] * PopTable![RELIGION])
       PopTable![TURN09] = CLng(PopTable![TURN09] * PopTable![HOSPITAL])
       If Current_Increase = 0 Then
          Current_Increase = PopTable![TURN09]
       End If
    End If
    
    PopTable.UPDATE

    ' perform addition of new increase
    TRIBESINFO.MoveFirst
    TRIBESINFO.Seek "=", TCLANNUMBER, TRIBES_POP_TRIBE
    Do Until Current_Increase = 0
       TRIBESINFO.Edit
       If TOldInactives < TOldActives Then
          If TOldInactives < TOldWarriors Then
             TRIBESINFO!INACTIVES = TRIBESINFO!INACTIVES + 1
             TOldInactives = TOldInactives + 1
             Current_Increase = Current_Increase - 1
          Else
             TRIBESINFO!WARRIORS = TRIBESINFO!WARRIORS + 1
             TOldWarriors = TOldWarriors + 1
             Current_Increase = Current_Increase - 1
          End If
       ElseIf TOldActives < TOldWarriors Then
             TRIBESINFO!ACTIVES = TRIBESINFO!ACTIVES + 1
             TOldActives = TOldActives + 1
             Current_Increase = Current_Increase - 1
       ElseIf TOldWarriors < TOldInactives Then
             TRIBESINFO!WARRIORS = TRIBESINFO!WARRIORS + 1
             TOldWarriors = TOldWarriors + 1
             Current_Increase = Current_Increase - 1
       ElseIf TOldWarriors = TOldActives Then
             If TOldWarriors = TOldInactives Then
                TRIBESINFO!WARRIORS = TRIBESINFO!WARRIORS + 1
                TOldWarriors = TOldWarriors + 1
                Current_Increase = Current_Increase - 1
             Else
                TRIBESINFO!INACTIVES = TRIBESINFO!INACTIVES + 1
                TOldInactives = TOldInactives + 1
                Current_Increase = Current_Increase - 1
             End If
       ElseIf TOldActives = TOldInactives Then
                TRIBESINFO!ACTIVES = TRIBESINFO!ACTIVES + 1
                TOldActives = TOldActives + 1
                Current_Increase = Current_Increase - 1
       Else
             TRIBESINFO!INACTIVES = TRIBESINFO!INACTIVES + 1
             TOldInactives = TOldInactives + 1
             Current_Increase = Current_Increase - 1
       End If
       TRIBESINFO.UPDATE
    Loop

ERR_Process_Population_Growth_CLOSE:
   Exit Function

ERR_Process_Population_Growth:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Population_Growth_CLOSE

End Function

Public Function Process_Slave_Growth()
Dim All_Slaves_Overseen As String
Dim AVAILABLE_SHACKLES As Long
Dim TOTAL_SLAVES As Long
Dim Total_People As Long

On Error GoTo ERR_Process_Slave_Growth
TRIBE_STATUS = "Process_Slave_Growth"

Call A150_Open_Tables("MODIFIERS")
MODTABLE.Seek "=", TTRIBENUMBER, "SLAVE INCREASE"
   
If MODTABLE.NoMatch Then
   Slave_Population_Increase = 0.007
Else
   MODTABLE.Edit
   Slave_Population_Increase = MODTABLE![AMOUNT]
End If
   
' Increase # of slaves - Should occur from about 72 slaves onwards
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
SlaveIncrease = CLng((TRIBESINFO![SLAVE] * Slave_Population_Increase))
NUMBER_OF_SLAVES = TRIBESINFO![SLAVE]
Total_People = TRIBESINFO![WARRIORS] + TRIBESINFO![ACTIVES] + TRIBESINFO![INACTIVES]

' if the tribe has the adoption research topic then all new growth should go into
' the tribe.  Will automatically add it into the Inactives.
RESEARCH_FOUND = "N"
Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "ADOPTION")
If RESEARCH_FOUND = "Y" Then
   TRIBESINFO.Edit
   TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] + SlaveIncrease
   TRIBESINFO.UPDATE
Else
   TRIBESINFO.Edit
   TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] + SlaveIncrease
   TRIBESINFO.UPDATE
End If
   
TOTAL_SLAVES = TRIBESINFO![SLAVE]

If TRIBESINFO![SLAVE] > 0 Then
    All_Slaves_Overseen = "NO"
   ' PERFORM CHECK TO SEE IF ALL SLAVES OVERSEEN ELSE ALLOW FOR RUNAWAYS
   ' TRIBES_PROCESSING TABLE
   ' GET THE NUMBER OF AVAILABLE SHACKLES
   Call A150_Open_Tables("IMPLEMENT_USAGE")
   ImplementUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, "SHACKLE"
   If Not ImplementUsage.NoMatch Then
      AVAILABLE_SHACKLES = ImplementUsage![total_available] - ImplementUsage![Number_Used]
      If AVAILABLE_SHACKLES > NUMBER_OF_SLAVES Then
         Call Update_Implement_Usage(TCLANNUMBER, GOODS_TRIBE, "SHACKLE", NUMBER_OF_SLAVES)
         NUMBER_OF_SLAVES = CLng(NUMBER_OF_SLAVES / 2)
      Else
         Call Update_Implement_Usage(TCLANNUMBER, GOODS_TRIBE, "SHACKLE", AVAILABLE_SHACKLES)
         NUMBER_OF_SLAVES = NUMBER_OF_SLAVES - AVAILABLE_SHACKLES
      End If
   End If
   
   If Tribes_Processing.BOF Then
       ' do nothing
   Else
       Tribes_Processing.MoveFirst
   End If
   Tribes_Processing.Seek "=", TTRIBENUMBER
   If Tribes_Processing.NoMatch Then
      ' No slaves have been overseen through activities
      ' REBEL/RUNAWAY
      ' calc what could be supervised
      ' calc howmany un-supervised
      If Tribes_Processing.NoMatch Then
         Tribes_Processing.AddNew
         Tribes_Processing![TRIBE] = TTRIBENUMBER
      Else
         Tribes_Processing.Edit
      End If
     
      If SLAVERY_LEVEL = 0 Then
         Tribes_Processing![All_Slaves_Overseen] = "N"
         Tribes_Processing![Number_Of_Slaves_Overseen] = 0
         All_Slaves_Overseen = "NO"
      ElseIf SLAVERY_LEVEL > 0 Then
         If ((Total_People / 10) * SLAVERY_LEVEL) < NUMBER_OF_SLAVES Then
            Tribes_Processing![All_Slaves_Overseen] = "N"
            All_Slaves_Overseen = "NO"
            Tribes_Processing![Number_Of_Slaves_Overseen] = (Total_People / 10) * SLAVERY_LEVEL
         Else
            Tribes_Processing![All_Slaves_Overseen] = "Y"
            All_Slaves_Overseen = "YES"
            Tribes_Processing![Number_Of_Slaves_Overseen] = TOTAL_SLAVES
         End If
      Else
         Tribes_Processing![All_Slaves_Overseen] = "N"
         All_Slaves_Overseen = "NO"
      End If
      Tribes_Processing.UPDATE
      
   ElseIf Tribes_Processing![All_Slaves_Overseen] = "Y" Then
        'DO NOTHING
        All_Slaves_Overseen = "YES"
   Else
       ' recalc
       Tribes_Processing.Edit
       If SLAVERY_LEVEL = 0 Then
          If Tribes_Processing![Warriors_Assigned] < (NUMBER_OF_SLAVES / 10) Then
             Tribes_Processing![All_Slaves_Overseen] = "N"
             Tribes_Processing![Number_Of_Slaves_Overseen] = Tribes_Processing![Warriors_Assigned] * 10
             All_Slaves_Overseen = "NO"
          Else
             Tribes_Processing![All_Slaves_Overseen] = "Y"
             Tribes_Processing![Number_Of_Slaves_Overseen] = TOTAL_SLAVES
             All_Slaves_Overseen = "YES"
          End If
       ElseIf ((Tribes_Processing![Warriors_Assigned] / 10) * SLAVERY_LEVEL) < NUMBER_OF_SLAVES Then
          Tribes_Processing![All_Slaves_Overseen] = "N"
          All_Slaves_Overseen = "NO"
       Else
          Tribes_Processing![All_Slaves_Overseen] = "Y"
          All_Slaves_Overseen = "YES"
       End If
       Tribes_Processing.UPDATE
      
   End If
   If All_Slaves_Overseen = "NO" Then
      'msg to say slaves could run
      Msg = "The number of slaves unsupervised in " & TTRIBENUMBER
      Msg = Msg & Chr(13) & Chr(10) & "is " & NUMBER_OF_SLAVES & " " & Chr(13) & Chr(10)
      Msg = Msg & "They should runaway or rebel."
      MsgBox (Msg)
   End If

End If

ERR_Process_Slave_Growth_CLOSE:
   Exit Function

ERR_Process_Slave_Growth:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Slave_Growth_CLOSE

End Function

Public Function Process_People_Eating()
On Error GoTo ERR_Process_People_Eating
TRIBE_STATUS = "Process_People_Eating"
Dim total_fish As Integer
Dim provs_eaten As Long

TProvsReq = 0
GoatsKilled = 0
CattleKilled = 0
CamelsKilled = 0
HorsesKilled = 0
provs_eaten = 0

TRIBESINFO.Edit
TActivesAvailable = TRIBESINFO![WARRIORS]
TActivesAvailable = TActivesAvailable + TRIBESINFO![ACTIVES]
TActivesAvailable = TActivesAvailable + TRIBESINFO![SLAVE]
TMouths = TActivesAvailable
TMouths = TMouths + TRIBESINFO![INACTIVES]

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "FINISHED", "PROVS"
If TRIBESGOODS.NoMatch Then
   TRIBESGOODS.AddNew
   TRIBESGOODS!CLAN = TCLANNUMBER
   TRIBESGOODS!TRIBE = TTRIBENUMBER
   TRIBESGOODS!ITEM_TYPE = "FINISHED"
   TRIBESGOODS!ITEM = "PROVS"
   TRIBESGOODS!ITEM_NUMBER = 0
   TRIBESGOODS.UPDATE
   TRIBESGOODS.MoveFirst
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ANIMAL", "HERDING DOG"
If Not TRIBESGOODS.NoMatch Then
   TMouths = TMouths + CLng(TRIBESGOODS![ITEM_NUMBER] / 2)
End If
   
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ANIMAL", "WARDOG"
If Not TRIBESGOODS.NoMatch Then
   TMouths = TMouths + CLng(TRIBESGOODS![ITEM_NUMBER] / 2)
End If

Msg = "Provs eaten: "

' UPDATE PROVS EATEN
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "MILK"
If TRIBESGOODS.NoMatch Then
   EXTRA_PROVS = 0
Else
   EXTRA_PROVS = TRIBESGOODS![ITEM_NUMBER] / 10
End If

If TMouths > 0 And EXTRA_PROVS > 0 Then
   If EXTRA_PROVS > TMouths Then
      Msg = Msg & TMouths & " milk, "
      EXTRA_PROVS = EXTRA_PROVS - TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths * 10
      TRIBESGOODS.UPDATE
      TMouths = 0

   Else
      Msg = Msg & EXTRA_PROVS & " milk, "
      TMouths = TMouths - EXTRA_PROVS
      EXTRA_PROVS = 0
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
   End If
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "BREAD"
If TRIBESGOODS.NoMatch Then
   'Do nothing
Else
   EXTRA_PROVS = EXTRA_PROVS + TRIBESGOODS![ITEM_NUMBER]
End If

If TMouths > 0 And EXTRA_PROVS > 0 Then
   If EXTRA_PROVS > TMouths Then
      Msg = Msg & TMouths & " bread, "
      EXTRA_PROVS = EXTRA_PROVS - TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   Else
      Msg = Msg & EXTRA_PROVS & " bread, "
      TMouths = TMouths - EXTRA_PROVS
      EXTRA_PROVS = 0
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
   End If
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "FISH"
If TRIBESGOODS.NoMatch Then
   'Do nothing
Else
   EXTRA_PROVS = EXTRA_PROVS + TRIBESGOODS![ITEM_NUMBER]
End If

If TMouths > 0 And EXTRA_PROVS > 0 Then
   If EXTRA_PROVS > TMouths Then
      Msg = Msg & TMouths & " fish, "
      EXTRA_PROVS = EXTRA_PROVS - TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   Else
      Msg = Msg & EXTRA_PROVS & " fish, "
      TMouths = TMouths - EXTRA_PROVS
      EXTRA_PROVS = 0
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
   End If
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER]
End If
  
If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " provs, "
   provs_eaten = 0
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "DRIED BREAD"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER]
End If
  

If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " dried bread, "
   provs_eaten = 0
End If
  
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "WAYBREAD"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER]
End If
  

If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " waybread, "
   provs_eaten = 0
End If
  
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "CHEESE"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER]
End If
  

If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TMouths
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " cheese, "
   provs_eaten = 0
End If
  
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "GRAIN"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER] / 40
End If
  

If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (TMouths * 40)
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " grain, "
   provs_eaten = 0
End If
  
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "GRAPE"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER] / 50
End If
  
If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (TMouths * 50)
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " grapes, "
   provs_eaten = 0
End If
  
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "GUT"
If TRIBESGOODS.NoMatch Then
   TOTAL_PROVS_AVAILABLE = 0
Else
   TOTAL_PROVS_AVAILABLE = TRIBESGOODS![ITEM_NUMBER] / 10
End If
  
If TMouths > 0 And TOTAL_PROVS_AVAILABLE > 0 Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      provs_eaten = TOTAL_PROVS_AVAILABLE
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
      TMouths = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      provs_eaten = TMouths
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - (TMouths * 10)
      TRIBESGOODS.UPDATE
      TMouths = 0
   End If
End If

If provs_eaten > 0 Then
   Msg = Msg & provs_eaten & " gut, "
   provs_eaten = 0
End If
  
TOTAL_PROVS_AVAILABLE = 0
  
'If TCLANNUMBER = "0330" Then
   Call Check_Turn_Output("", Msg, "", 0, "NO")
'End If

If TOTAL_PROVS_AVAILABLE < TMouths Then
   If TOTAL_PROVS_AVAILABLE < TMouths Then
      TProvsReq = TMouths - TOTAL_PROVS_AVAILABLE
   Else
      TProvsReq = 0
   End If

   If TProvsReq > 0 Then
      TFINISHED = "NO"
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "GOAT"
      If Not TRIBESGOODS.NoMatch Then
         Do While TFINISHED = "NO"
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
               TRIBESGOODS.UPDATE
               If TRIBESGOODS![ITEM_NUMBER] < 1 Then
                  TFINISHED = "YES"
               End If
               TProvsReq = TProvsReq - 4
               If TProvsReq < 1 Then
                  TFINISHED = "YES"
               End If
               GoatsKilled = GoatsKilled + 1
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + 4
               TRIBESGOODS.UPDATE
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "GOAT"
            Else
               TFINISHED = "YES"
            End If
         Loop
         If GoatsKilled > 0 Then
            Call Check_Turn_Output(",", " goats killed ", "", GoatsKilled, "NO")
         End If
      End If
   End If
      
   If TProvsReq > 0 Then
      TFINISHED = "NO"
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "CATTLE"
      If Not TRIBESGOODS.NoMatch Then
         Do While TFINISHED = "NO"
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
               TRIBESGOODS.UPDATE
               If TRIBESGOODS![ITEM_NUMBER] < 1 Then
                  TFINISHED = "YES"
               End If
               TProvsReq = TProvsReq - 20
               If TProvsReq < 1 Then
                  TFINISHED = "YES"
               End If
               CattleKilled = CattleKilled + 1
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + 20
               TRIBESGOODS.UPDATE
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "CATTLE"
            Else
               TFINISHED = "YES"
            End If
         Loop
         If CattleKilled > 0 Then
            Call Check_Turn_Output(",", " cattle killed ", "", CattleKilled, "NO")
         End If
      End If
   End If
 
   If TProvsReq > 0 Then
      TFINISHED = "NO"
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "CAMEL"
      If Not TRIBESGOODS.NoMatch Then
         Do While TFINISHED = "NO"
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
               TRIBESGOODS.UPDATE
               If TRIBESGOODS![ITEM_NUMBER] < 1 Then
                  TFINISHED = "YES"
               End If
               TProvsReq = TProvsReq - 30
               If TProvsReq < 1 Then
                  TFINISHED = "YES"
               End If
               CamelsKilled = CamelsKilled + 1
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + 30
               TRIBESGOODS.UPDATE
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "CAMEL"
            Else
               TFINISHED = "YES"
            End If
         Loop
         If CamelsKilled > 0 Then
            Call Check_Turn_Output(",", " camels killed ", "", CamelsKilled, "NO")
         End If
      End If
   End If

   If TProvsReq > 0 Then
      TFINISHED = "NO"
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "HORSE"
      If Not TRIBESGOODS.NoMatch Then
         Do While TFINISHED = "NO"
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
               TRIBESGOODS.UPDATE
               If TRIBESGOODS![ITEM_NUMBER] < 1 Then
                  TFINISHED = "YES"
               End If
               TProvsReq = TProvsReq - 30
               If TProvsReq < 1 Then
                  TFINISHED = "YES"
               End If
               HorsesKilled = HorsesKilled + 1
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + 30
               TRIBESGOODS.UPDATE
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "HORSE"
            Else
               TFINISHED = "YES"
            End If
         Loop
         If HorsesKilled > 0 Then
            Call Check_Turn_Output(",", " horses killed ", "", HorsesKilled, "NO")
         End If
      End If
   End If
End If
 
Set PROVS_AVAIL_TABLE = TVDBGM.OpenRecordset("Provs_Availability")
PROVS_AVAIL_TABLE.index = "PRIMARYKEY"
If Not PROVS_AVAIL_TABLE.EOF Then
   PROVS_AVAIL_TABLE.MoveFirst
End If
PROVS_AVAIL_TABLE.Seek "=", TTRIBENUMBER
  
If PROVS_AVAIL_TABLE.NoMatch Then
   PROVS_AVAIL_TABLE.AddNew
   PROVS_AVAIL_TABLE![TRIBE] = TTRIBENUMBER
   PROVS_AVAIL_TABLE![WARNED] = "N"
   PROVS_AVAIL_TABLE![POP_LOSS] = 0
   PROVS_AVAIL_TABLE.UPDATE
   PROVS_AVAIL_TABLE.MoveFirst
   PROVS_AVAIL_TABLE.Seek "=", TTRIBENUMBER
End If
  
POPULATION_STARVED = 0

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
If TMouths > 0 Then
  If TRIBESGOODS.NoMatch Then
     Call Check_Turn_Output(",", " There are insufficient Provs to feed the people in this group next turn.", "", 0, "NO")
     Msg = GOODS_TRIBE & " IS MISSING " & TMouths & " PROVS TO FEED PEOPLE"
     MsgBox (Msg)
     If PROVS_AVAIL_TABLE![WARNED] = "Y" Then
        POPULATION_STARVED = PROVS_AVAIL_TABLE![POP_LOSS] + 10
        POPULATION_STARVED = (POPULATION_STARVED / 100) * TMouths
     End If
     
  ElseIf TMouths > TRIBESGOODS![ITEM_NUMBER] Then
     Call Check_Turn_Output(",", " There are insufficient Provs to feed the people in this group next turn.", "", 0, "NO")
     Msg = GOODS_TRIBE & "IS MISSING " & (TMouths - TRIBESGOODS![ITEM_NUMBER]) & "PROVS TO FEED PEOPLE"
     MsgBox (Msg)
     If PROVS_AVAIL_TABLE![WARNED] = "Y" Then
        POPULATION_STARVED = PROVS_AVAIL_TABLE![POP_LOSS] + 10
        POPULATION_STARVED = (POPULATION_STARVED / 100) * TMouths
     End If
     TRIBESGOODS.Edit
     TRIBESGOODS![ITEM_NUMBER] = 0
     TRIBESGOODS.UPDATE
  Else
     TRIBESGOODS.Edit
     TRIBESGOODS![ITEM_NUMBER] = (TRIBESGOODS![ITEM_NUMBER] - TMouths)
     TRIBESGOODS.UPDATE
     PROVS_AVAIL_TABLE.Edit
     PROVS_AVAIL_TABLE![WARNED] = "N"
     PROVS_AVAIL_TABLE![POP_LOSS] = 0
     PROVS_AVAIL_TABLE.UPDATE
  End If
End If

If POPULATION_STARVED > 0 Then
   TRIBESINFO.Edit
   TOTAL_POPULATION = TRIBESINFO![WARRIORS]
   TOTAL_POPULATION = TOTAL_POPULATION + TRIBESINFO![ACTIVES]
   TOTAL_POPULATION = TOTAL_POPULATION + TRIBESINFO![INACTIVES]
   TOTAL_POPULATION = TOTAL_POPULATION + TRIBESINFO![SLAVE]
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ANIMAL", "HERDING DOG"
   If Not TRIBESGOODS.NoMatch Then
      If POPULATION_STARVED > TRIBESGOODS![ITEM_NUMBER] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESGOODS![ITEM_NUMBER]
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = 0
         TRIBESGOODS.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - POPULATION_STARVED
         TRIBESGOODS.UPDATE
      End If
   End If
 
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ANIMAL", "WARDOG"
   If Not TRIBESGOODS.NoMatch Then
      If POPULATION_STARVED > TRIBESGOODS![ITEM_NUMBER] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESGOODS![ITEM_NUMBER]
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = 0
         TRIBESGOODS.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - POPULATION_STARVED
         TRIBESGOODS.UPDATE
      End If
   End If
   If TRIBESINFO![SLAVE] > 0 Then
      If POPULATION_STARVED > TRIBESINFO![SLAVE] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESINFO![SLAVE]
         TRIBESINFO.Edit
         TRIBESINFO![SLAVE] = 0
         TRIBESINFO.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESINFO.Edit
         TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - POPULATION_STARVED
         TRIBESINFO.UPDATE
      End If
      If POPULATION_STARVED > TRIBESINFO![INACTIVES] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESINFO![INACTIVES]
         TRIBESINFO.Edit
         TRIBESINFO![INACTIVES] = 0
         TRIBESINFO.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESINFO.Edit
         TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] - POPULATION_STARVED
         TRIBESINFO.UPDATE
      End If
      If POPULATION_STARVED > TRIBESINFO![ACTIVES] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESINFO![ACTIVES]
         TRIBESINFO.Edit
         TRIBESINFO![ACTIVES] = 0
         TRIBESINFO.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESINFO.Edit
         TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] - POPULATION_STARVED
         TRIBESINFO.UPDATE
      End If
      If POPULATION_STARVED > TRIBESINFO![WARRIORS] Then
         POPULATION_STARVED = POPULATION_STARVED - TRIBESINFO![WARRIORS]
         TRIBESINFO.Edit
         TRIBESINFO![WARRIORS] = 0
         TRIBESINFO.UPDATE
      Else
         POPULATION_STARVED = 0
         TRIBESINFO.Edit
         TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - POPULATION_STARVED
         TRIBESINFO.UPDATE
      End If
   End If
End If
' Warn about low provs levels & record the warning
' How many months
' TMouths = "how many this turn"
' TribesGoods![ITEM_NUMBER] = "how many provs left at end of turn"

If TMouths > 0 Then
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "PROVS"
   If TRIBESGOODS.NoMatch Then
      ' output warning
      Call Check_Turn_Output(",", " Provs on hand will not feed people ", "", 0, "NO")
      ' update table
      PROVS_AVAIL_TABLE.MoveFirst
      PROVS_AVAIL_TABLE.Seek "=", TTRIBENUMBER
      If PROVS_AVAIL_TABLE![WARNED] = "Y" Then
         PROVS_AVAIL_TABLE.Edit
         PROVS_AVAIL_TABLE![WARNED] = "Y"
         PROVS_AVAIL_TABLE![POP_LOSS] = PROVS_AVAIL_TABLE![POP_LOSS] + 10
         PROVS_AVAIL_TABLE.UPDATE
      Else
         PROVS_AVAIL_TABLE.Edit
         PROVS_AVAIL_TABLE![WARNED] = "Y"
         PROVS_AVAIL_TABLE![POP_LOSS] = 0
         PROVS_AVAIL_TABLE.UPDATE
      End If
   ElseIf TMouths > TRIBESGOODS![ITEM_NUMBER] Then
      ' output warning
      Call Check_Turn_Output(",", " Provs on hand will not feed people ", "", TRIBESGOODS![ITEM_NUMBER], "NO")
      ' update table
      PROVS_AVAIL_TABLE.MoveFirst
      PROVS_AVAIL_TABLE.Seek "=", TTRIBENUMBER
      If PROVS_AVAIL_TABLE![WARNED] = "Y" Then
         PROVS_AVAIL_TABLE.Edit
         PROVS_AVAIL_TABLE![WARNED] = "Y"
         PROVS_AVAIL_TABLE![POP_LOSS] = PROVS_AVAIL_TABLE![POP_LOSS] + 10
         PROVS_AVAIL_TABLE.UPDATE
      Else
         PROVS_AVAIL_TABLE.Edit
         PROVS_AVAIL_TABLE![WARNED] = "Y"
         PROVS_AVAIL_TABLE![POP_LOSS] = 0
         PROVS_AVAIL_TABLE.UPDATE
      End If
   Else
      ' update table
      PROVS_AVAIL_TABLE.Edit
      PROVS_AVAIL_TABLE![WARNED] = "N"
      PROVS_AVAIL_TABLE![POP_LOSS] = 0
      PROVS_AVAIL_TABLE.UPDATE
   End If
End If

ERR_Process_People_Eating_CLOSE:
   Exit Function

ERR_Process_People_Eating:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_People_Eating_CLOSE

End Function


Public Function Process_Implement_Usage(ACTIVITY, ACTIVITY_TYPE, Variable_Increased, Reduce_Actives)
On Error GoTo ERR_Process_Implement_Usage
TRIBE_STATUS = "Process_Implement_Usage"
   
ImplementsTable.index = "ACTIVITY"
ImplementsTable.MoveFirst
ImplementsTable.Seek "=", ACTIVITY, ACTIVITY_TYPE
If ImplementsTable.NoMatch Then
   Exit Function
End If

IMPLEMENT = ImplementsTable![IMPLEMENT]
IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
   
Do While (ImplementsTable![ACTIVITY] = ACTIVITY)
   'check for goods tribe
   PROCESSITEMS.MoveFirst
   PROCESSITEMS.Seek "=", TTRIBENUMBER, ACTIVITY, TItem, IMPLEMENT
   If Not PROCESSITEMS.NoMatch Then
      If TActives > 0 Then
         TImplement = PROCESSITEMS![QUANTITY]
         If TImplement > 0 Then
            total_available = PROCESSITEMS![QUANTITY]
            If IMPLEMENT = "TRAP" Then
               If (TImplement / TRAPS_TO_USE) > TActives Then
                  Variable_Increased = (Variable_Increased + ((TActives * TRAPS_TO_USE) _
                  * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = 0
                  End If
               Else
                  Variable_Increased = (Variable_Increased + (TImplement * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = TActives - (TImplement / TRAPS_TO_USE)
                  End If
               End If
            ElseIf IMPLEMENT = "SNARE" Then
               If (TImplement / SNARES_TO_USE) > TActives Then
                  Variable_Increased = (Variable_Increased + ((TActives * SNARES_TO_USE) _
                  * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = 0
                  End If
               Else
                  Variable_Increased = (Variable_Increased + (TImplement * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = TActives - (TImplement / SNARES_TO_USE)
                  End If
               End If
            ElseIf IMPLEMENT = "MINING LADDER" Then
                If (TActives / 5) >= TImplement Then
                  Variable_Increased = Variable_Increased + TActives
               Else
                  Variable_Increased = Variable_Increased + (TImplement * 5)
               End If
            ElseIf IMPLEMENT = "ORE CART" Then
                If (TActives / 5) >= TImplement Then
                  Variable_Increased = Variable_Increased + TActives
               Else
                  Variable_Increased = Variable_Increased + (TImplement * 5)
               End If
            Else
               If TImplement >= TActives Then
                  Variable_Increased = (Variable_Increased + (TActives * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = 0
                  End If
               Else
                  Variable_Increased = (Variable_Increased + (TImplement * IMPLEMENT_MODIFIER))
                  If Reduce_Actives = "YES" Then
                     TActives = TActives - TImplement
                  End If
               End If
            End If
         End If
      End If
   End If
   ImplementsTable.MoveNext
   If ImplementsTable.EOF Then
      Exit Do
   ElseIf Not (ImplementsTable![ITEM] = ACTIVITY_TYPE) Then
      IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
      Exit Do
   Else
      IMPLEMENT = ImplementsTable![IMPLEMENT]
      IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
   End If
Loop
       
ImplementsTable.index = "PRIMARYKEY"
ImplementsTable.MoveFirst

ERR_Process_Implement_Usage_CLOSE:
   Exit Function

ERR_Process_Implement_Usage:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Implement_Usage_CLOSE

End Function

Public Function PERFORM_COMMON(Joint_Check, Rest, Table_Update, Table_Update_Loops As Long, Order As String)
'*===============================================================================*'
'*****                          MAINTENANCE LOG                              *****'
'*-------------------------------------------------------------------------------*'
'**  VARIABLE    *  DESCRIPTION                                                 **'
'*-------------------------------------------------------------------------------*'
'** Joint_Check  *  Check if joint activity                                     **'
'** Rest         *  Reset number occurs                                         **'
'** Verify       *  Verify Quantities                                           **'
'** Verify_Loops *  Number of quantities to check                               **'
'** Table_Update *  Update tables                                               **'
'** Table_Update_Loops *  Number of goods to update                             **'
'** Order        *  order for printing                                          **'
'**                                                                             **'
'*===============================================================================*'On Error GoTo ERR_COMMON
TRIBE_STATUS = "PERFORM_COMMON"
  
TempOutput = " (using"
  
  
If Joint_Check = "Y" Then
   If TJoint = "Y" Then
      TActives = (TActives * (10 / (10 + SKILL_SHORTAGE)))
   End If
End If

If Rest = "Y" Then
   If TPeople > 0 Then
      TNUMOCCURS = TActives / TPeople
      ActivesNeeded = TNUMOCCURS * TPeople
      NumItemsMade = TNUMOCCURS * TNumItems
   Else
      TNUMOCCURS = 0
      ActivesNeeded = 0
      NumItemsMade = 0
   End If
    
      If TActivesAvailable < ActivesNeeded Then
         TNUMOCCURS = CLng(TActivesAvailable / TPeople)
      End If
   
      ModifyTable = "Y"
End If

Call Verify_Quantities

If Table_Update = "Y" Then
   If ModifyTable = "Y" Then
      Index1 = 1
      Do Until Index1 > Table_Update_Loops
         If TGoods(Index1) = "EMPTY" Then
            Index1 = 20
         Else
            BRACKET = InStr(TGoods(Index1), "(")
            If BRACKET > 0 Then
               ITEM = Left(TGoods(Index1), (BRACKET - 1))
            Else
               ITEM = TGoods(Index1)
            End If
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "SUBTRACT", (TNUMOCCURS * TQuantity(Index1)))
            If Index1 = 1 Then
               TempOutput = TempOutput & " " & (TNUMOCCURS * TQuantity(Index1)) & " " & StrConv(ITEM, vbProperCase)
            Else
               TempOutput = TempOutput & ", " & (TNUMOCCURS * TQuantity(Index1)) & " " & StrConv(ITEM, vbProperCase)
            End If
         End If
         Index1 = Index1 + 1
      Loop
      
      BRACKET = InStr(TItem, "(")
      If BRACKET > 0 Then
         ITEM = Left(TItem, (BRACKET - 1))
      Else
         ITEM = TItem
      End If
      If TItem_Produced = "Forget" Then
         'do nothing
      Else
         ITEM = TItem_Produced
      End If
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, ITEM, "ADD", NumItemsMade)
   End If
  
   ' update output line
   If ModifyTable = "Y" Then
      If Order = "ANI" Then
         Call UPDATE_TURNACTOUTPUT("ANI")
      ElseIf Order = "ANSI" Then
         Call UPDATE_TURNACTOUTPUT("ANSI")
      ElseIf Order = "ASNI" Then
         Call UPDATE_TURNACTOUTPUT("ASNI")
      Else
         Call UPDATE_TURNACTOUTPUT("NO")
      End If
      TempOutput = TempOutput & ")"
      TurnActOutPut = TurnActOutPut & TempOutput
   End If
End If

ERR_COMMON_CLOSE:
   Exit Function

ERR_COMMON:
If (Err = 3021) Or (Err = 3022) Then
   ' 3021 - No current record
   ' 3022 - Duplicate Record
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_COMMON_CLOSE
End If

End Function

Public Function PERFORM_GATHERING()
On Error GoTo ERR_PERFORM_GATHERING
TRIBE_STATUS = "PERFORM_GATHERING"

TOTAL_SAND = 0
FODDER_GATHER = 0

If TActivity = "Sand Gathering" Or TItem = "SAND" Then
   Call Process_Implement_Usage("GATHERING", "SAND", TOTAL_SAND, "NO")
   
   TOTAL_SAND = (TActives * 20)
       
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SAND", "ADD", TOTAL_SAND)

   Call Check_Turn_Output(", ", " effective people gather", " Sand", TOTAL_SAND, "YES")
       
ElseIf TItem = "FODDER" Then
   Select Case TRIBES_TERRAIN
   Case "GRASSY HILLS", "PRAIRIE"
      Call Process_Implement_Usage("FORAGING", "ALL", TActives, "NO")
     
      FODDER_GATHER = TActives * 50
      
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "FODDER", "ADD", FODDER_GATHER)
      
      If Len(TurnActOutPut) > 20 Then
         If Right(TurnActOutPut, 1) = " " Or Right(TurnActOutPut, 2) = ") " Or IsLetter(Right(TurnActOutPut, 1)) Then
            TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people gather " & FODDER_GATHER & " Fodder, "
         Else
            TurnActOutPut = TurnActOutPut & TActives & " effective people gather " & FODDER_GATHER & " Fodder, "
         End If
      Else
         TurnActOutPut = TurnActOutPut & TActives & " effectivepeople gather " & FODDER_GATHER & " Fodder, "
      End If
   Case Else
      If Len(TurnActOutPut) > 20 Then
         If Right(TurnActOutPut, 1) = " " Or Right(TurnActOutPut, 2) = ") " Or IsLetter(Right(TurnActOutPut, 1)) Then
            TurnActOutPut = TurnActOutPut & ", Invalid terrain for foraging "
         Else
            TurnActOutPut = TurnActOutPut & " Invalid terrain for foraging "
         End If
      Else
         TurnActOutPut = TurnActOutPut & " Invalid terrain for foraging "
      End If
   End Select

ElseIf TItem = "WATER" Then
   LIQUID_STORAGE = 0
   LIQUID_ONHAND = 0
   LIQUID_STORAGE_AVAILABLE = 0
   
   Call DETERMINE_LIQUID_STORAGE
   
   Call DETERMINE_LIQUID_ONHAND

   If LIQUID_ONHAND > LIQUID_STORAGE Then
      LIQUID_STORAGE_AVAILABLE = 0
   Else
      LIQUID_STORAGE_AVAILABLE = LIQUID_STORAGE - LIQUID_ONHAND
   End If
   
   If LIQUID_STORAGE_AVAILABLE > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WATER", "ADD", LIQUID_STORAGE_AVAILABLE)
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "FINISHED", "WATER"
      If Len(TurnActOutPut) > 20 Then
         If Right(TurnActOutPut, 1) = " " Or Right(TurnActOutPut, 2) = ") " Then
            TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people gather " & LIQUID_STORAGE_AVAILABLE & " Water, "
         Else
            TurnActOutPut = TurnActOutPut & TActives & " effective people gather " & LIQUID_STORAGE_AVAILABLE & " Water, "
         End If
      Else
         TurnActOutPut = TurnActOutPut & TActives & " effective people gather " & LIQUID_STORAGE_AVAILABLE & " Water, "
      End If
   Else
      TurnActOutPut = TurnActOutPut & " No room left storing water, "
   End If
   
End If

ERR_PERFORM_GATHERING_CLOSE:
   Exit Function

ERR_PERFORM_GATHERING:
   Call A999_ERROR_HANDLING
   Resume ERR_PERFORM_GATHERING_CLOSE

End Function

Public Function DETERMINE_LIQUID_STORAGE()
On Error GoTo ERR_DETERMINE_LIQUID_STORAGE
TRIBE_STATUS = "DETERMINE_LIQUID_STORAGE"
   
LIQUID_STORAGE = 0

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BARREL")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 100)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BEAKERS")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BLADDER")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 10)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "EWER")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 20)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "JAR")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 50)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "KEG")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 400)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "URN")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 150)
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WATERTANK")
   
If Num_Goods > 0 Then
   LIQUID_STORAGE = LIQUID_STORAGE + (Num_Goods * 1000)
End If

ERR_DETERMINE_LIQUID_STORAGE_CLOSE:
   Exit Function

ERR_DETERMINE_LIQUID_STORAGE:
   Call A999_ERROR_HANDLING
   Resume ERR_DETERMINE_LIQUID_STORAGE_CLOSE

End Function

Public Function DETERMINE_LIQUID_ONHAND()
On Error GoTo ERR_DETERMINE_LIQUID_ONHAND
TRIBE_STATUS = "DETERMINE_LIQUID_ONHAND"

LIQUID_ONHAND = 0

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "ALE")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "BRANDY")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "COAL TAR")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "LINSEED OIL")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "MILK")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "MEAD")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "OIL")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "RUM")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "VODKA")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WATER")
   
If Num_Goods > 0 Then
   WATER_ONHAND = TRIBESGOODS![ITEM_NUMBER]
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WINE")
   
If Num_Goods > 0 Then
   LIQUID_ONHAND = LIQUID_ONHAND + Num_Goods
End If

ERR_DETERMINE_LIQUID_ONHAND_CLOSE:
   Exit Function

ERR_DETERMINE_LIQUID_ONHAND:
   Call A999_ERROR_HANDLING
   Resume ERR_DETERMINE_LIQUID_ONHAND_CLOSE

End Function


Public Function Populate_Tribes_Goods_Usage_Table()
On Error GoTo ERR_POPULATE_GOODS
TRIBE_STATUS = "Populate_Tribes_Goods_Usage_Table"

' This is used in the Clean_Up_and_Reset process.
' This occurs at the start of processing.

TRIBESGOODS.index = "CLANTRIBE"
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE

If Not TRIBESGOODS.EOF Then
  Do Until TRIBESGOODS![TRIBE] <> GOODS_TRIBE
     ' READ THE GOODS TABLE
  
    TribesGoodsUsage.AddNew
    TribesGoodsUsage![CLAN] = TCLANNUMBER
    TribesGoodsUsage![TRIBE] = GOODS_TRIBE
    TribesGoodsUsage![ITEM] = TRIBESGOODS![ITEM]
    TribesGoodsUsage![total_available] = TRIBESGOODS![ITEM_NUMBER]
    TribesGoodsUsage![Number_Used] = 0
    TribesGoodsUsage.UPDATE
     
    TRIBESGOODS.MoveNext
    If TRIBESGOODS.EOF Then
       Exit Do
    End If
  Loop
End If
TRIBESGOODS.index = "PRIMARYKEY"

ERR_POP_GOODS_CLOSE:
   Exit Function


ERR_POPULATE_GOODS:
If (Err = 3021) Or (Err = 3022) Or (Err = 3058) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   TRIBESGOODS.index = "PRIMARYKEY"
   Resume ERR_POP_GOODS_CLOSE
End If


End Function

Public Function Process_Tribes_Goods_Usage(QUESTION, ACTIVITY, ITEM)
On Error GoTo ERR_Process_Tribes_Goods_Usage
TRIBE_STATUS = "Process_Tribes_Goods_Usage"
     
TribesGoodsUsage.MoveFirst
TribesGoodsUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, ITEM
       
If Not TribesGoodsUsage.NoMatch Then
   TImplement = InputBox(QUESTION, ACTIVITY, "0")
   If TImplement > 0 Then
      total_available = ImplementUsage![total_available] - ImplementUsage![Number_Used]
      If TImplement > total_available Then
         TImplement = total_available
      End If
      If TImplement > TActives Then
         TActives = (TActives + (TActives * ImplementUsage![Modifier]))
         ImplementUsage.Edit
         ImplementUsage![Number_Used] = ImplementUsage![Number_Used] + TActives
         ImplementUsage.UPDATE
      Else
         TActives = (TActives + (TImplement * ImplementUsage![Modifier]))
         ImplementUsage.Edit
         ImplementUsage![Number_Used] = ImplementUsage![Number_Used] + TImplement
         ImplementUsage.UPDATE
      End If
   End If
End If

ERR_Process_Tribes_Goods_Usage_CLOSE:
   Exit Function

ERR_Process_Tribes_Goods_Usage:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Tribes_Goods_Usage_CLOSE

End Function


Public Function Perform_Politics_Activities(GL_LEVEL)
On Error GoTo ERR_Perform_Politics_Activities
TRIBE_STATUS = "Perform_Politics_Activities"

' this needs to occur once for each tribe.  Similiar to herding.

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_HEX
CURRENT_TERRAIN = hexmaptable![TERRAIN]

HEXMAPPOLITICS.MoveFirst
HEXMAPPOLITICS.Seek "=", CURRENT_HEX

If HEXMAPPOLITICS.NoMatch Then
   'MSG = "HEX OF TRIBE NOT FOUND"
   'MsgBox (MSG)
   CURRENT_HEX_POP = 0
   CURRENT_HEX_PAC_LEV = 0
Else
   If IsNull(HEXMAPPOLITICS![POPULATION]) Then
      CURRENT_HEX_POP = 0
   Else
      CURRENT_HEX_POP = HEXMAPPOLITICS![POPULATION]
   End If
   If IsNull(HEXMAPPOLITICS![PACIFICATION_LEVEL]) Then
      CURRENT_HEX_PAC_LEV = 0
   Else
      CURRENT_HEX_PAC_LEV = HEXMAPPOLITICS![PACIFICATION_LEVEL]
   End If
End If

If CURRENT_HEX_PAC_LEV >= 1 Then
   AvailableSerfs = CLng(CURRENT_HEX_POP * ((CURRENT_HEX_PAC_LEV * 2) / 100))
   ' Get stones
   If GL_LEVEL = "GL0" Then
       STONES_QUARRIED = CLng(AvailableSerfs * 2.5)
   ElseIf GL_LEVEL = "GL1" Then
       STONES_QUARRIED = CLng((AvailableSerfs * 2.5) / 2)
   ElseIf GL_LEVEL = "GL2" Then
       STONES_QUARRIED = CLng(((AvailableSerfs * 2.5) / 2) / 2)
   ElseIf GL_LEVEL = "GL3" Then
       STONES_QUARRIED = CLng((((AvailableSerfs * 2.5) / 2) / 2) / 2)
   ElseIf GL_LEVEL = "GL4" Then
       STONES_QUARRIED = CLng(((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2)
   ElseIf GL_LEVEL = "GL5" Then
       STONES_QUARRIED = CLng((((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2) / 2)
   End If
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "STONE", "ADD", STONES_QUARRIED)
   OutLine = OutLine & "Stones - " & STONES_QUARRIED & ","
         
   ' Get minerals
   HEXMAPMINERALS.MoveFirst
   HEXMAPMINERALS.Seek "=", CURRENT_HEX
   
   If Not HEXMAPMINERALS.NoMatch Then
       Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Geologists")
       If RESEARCH_FOUND = "Y" Then
           If Not IsNull(HEXMAPMINERALS![ORE_TYPE]) And Not (HEXMAPMINERALS![ORE_TYPE]) = "NONE" Then
               If GL_LEVEL = "GL0" Then
                   TNEWORE = CLng(AvailableSerfs * 2.5)
               ElseIf GL_LEVEL = "GL1" Then
                   TNEWORE = CLng((AvailableSerfs * 2.5) / 2)
               ElseIf GL_LEVEL = "GL2" Then
                   TNEWORE = CLng(((AvailableSerfs * 2.5) / 2) / 2)
               ElseIf GL_LEVEL = "GL3" Then
                   TNEWORE = CLng((((AvailableSerfs * 2.5) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL4" Then
                   TNEWORE = CLng(((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL5" Then
                   TNEWORE = CLng((((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2) / 2)
               End If
               Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, HEXMAPMINERALS![ORE_TYPE], "ADD", TNEWORE)
           ElseIf Not IsNull(HEXMAPMINERALS![SECOND_ORE]) And Not (HEXMAPMINERALS![SECOND_ORE]) = "NONE" Then
               If GL_LEVEL = "GL0" Then
                   TNEWORE = CLng(AvailableSerfs * 2)
               ElseIf GL_LEVEL = "GL1" Then
                   TNEWORE = CLng((AvailableSerfs * 2) / 2)
               ElseIf GL_LEVEL = "GL2" Then
                   TNEWORE = CLng(((AvailableSerfs * 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL3" Then
                   TNEWORE = CLng((((AvailableSerfs * 2) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL4" Then
                   TNEWORE = CLng(((((AvailableSerfs * 2) / 2) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL5" Then
                   TNEWORE = CLng((((((AvailableSerfs * 2) / 2) / 2) / 2) / 2) / 2)
               End If
               Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, HEXMAPMINERALS![SECOND_ORE], "ADD", TNEWORE)
           ElseIf Not IsNull(HEXMAPMINERALS![THIRD_ORE]) And Not (HEXMAPMINERALS![THIRD_ORE]) = "NONE" Then
               If GL_LEVEL = "GL0" Then
                   TNEWORE = CLng(AvailableSerfs * 1.5)
               ElseIf GL_LEVEL = "GL1" Then
                   TNEWORE = CLng((AvailableSerfs * 1.5) / 2)
               ElseIf GL_LEVEL = "GL2" Then
                   TNEWORE = CLng(((AvailableSerfs * 1.5) / 2) / 2)
               ElseIf GL_LEVEL = "GL3" Then
                   TNEWORE = CLng((((AvailableSerfs * 1.5) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL4" Then
                   TNEWORE = CLng(((((AvailableSerfs * 1.5) / 2) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL5" Then
                   TNEWORE = CLng((((((AvailableSerfs * 1.5) / 2) / 2) / 2) / 2) / 2)
               End If
               Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, HEXMAPMINERALS![THIRD_ORE], "ADD", TNEWORE)
           ElseIf Not IsNull(HEXMAPMINERALS![FORTH_ORE]) And Not (HEXMAPMINERALS![FORTH_ORE]) = "NONE" Then
               If GL_LEVEL = "GL0" Then
                   TNEWORE = CLng(AvailableSerfs * 1)
               ElseIf GL_LEVEL = "GL1" Then
                   TNEWORE = CLng((AvailableSerfs * 1) / 2)
               ElseIf GL_LEVEL = "GL2" Then
                   TNEWORE = CLng(((AvailableSerfs * 1) / 2) / 2)
               ElseIf GL_LEVEL = "GL3" Then
                   TNEWORE = CLng((((AvailableSerfs * 1) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL4" Then
                   TNEWORE = CLng(((((AvailableSerfs * 1) / 2) / 2) / 2) / 2)
               ElseIf GL_LEVEL = "GL5" Then
                   TNEWORE = CLng((((((AvailableSerfs * 1) / 2) / 2) / 2) / 2) / 2)
               End If
               Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, HEXMAPMINERALS![FORTH_ORE], "ADD", TNEWORE)
           End If
      Else
           If GL_LEVEL = "GL0" Then
               TNEWORE = CLng(AvailableSerfs * 2.5)
           ElseIf GL_LEVEL = "GL1" Then
               TNEWORE = CLng((AvailableSerfs * 2.5) / 2)
           ElseIf GL_LEVEL = "GL2" Then
               TNEWORE = CLng(((AvailableSerfs * 2.5) / 2) / 2)
           ElseIf GL_LEVEL = "GL3" Then
               TNEWORE = CLng((((AvailableSerfs * 2.5) / 2) / 2) / 2)
           ElseIf GL_LEVEL = "GL4" Then
               TNEWORE = CLng(((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2)
           ElseIf GL_LEVEL = "GL5" Then
               TNEWORE = CLng((((((AvailableSerfs * 2.5) / 2) / 2) / 2) / 2) / 2)
           End If
           Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, HEXMAPMINERALS![ORE_TYPE], "ADD", TNEWORE)
      End If
      OutLine = OutLine & " " & HEXMAPMINERALS![ORE_TYPE] & " - " & TNEWORE & ","
   End If
   
   ' Get logs & bark
   Select Case CURRENT_TERRAIN
   Case "CONIFER HILLS"
      FORESTRY_OK = "Y"
   Case "DECIDUOUS"
      FORESTRY_OK = "Y"
   Case "DECIDUOUS FLAT"
      FORESTRY_OK = "Y"
   Case "DECIDUOUS FOREST"
      FORESTRY_OK = "Y"
   Case "DECIDUOUS HILLS"
      FORESTRY_OK = "Y"
   Case "HARDWOOD FOREST"
      FORESTRY_OK = "Y"
   Case "JUNGLE"
      FORESTRY_OK = "Y"
   Case "JUNGLE HILLS"
      FORESTRY_OK = "Y"
   Case "LOW CONIFER MOUNTAINS"
      FORESTRY_OK = "Y"
   Case "LOW CONIFER MT"
      FORESTRY_OK = "Y"
   Case "LOW JUNGLE MOUNTAINS"
      FORESTRY_OK = "Y"
   Case "LOW JUNGLE MT"
      FORESTRY_OK = "Y"
   Case "MANGROVE SWAMP"
      FORESTRY_OK = "Y"
   Case Else
      FORESTRY_OK = "N"
   End Select
       
   If FORESTRY_OK = "Y" Then
       If GL_LEVEL = "GL0" Then
           TLogs = CLng(AvailableSerfs * 2)
       ElseIf GL_LEVEL = "GL1" Then
           TLogs = CLng((AvailableSerfs * 2) / 2)
       ElseIf GL_LEVEL = "GL2" Then
           TLogs = CLng(((AvailableSerfs * 2) / 2) / 2)
       ElseIf GL_LEVEL = "GL3" Then
           TLogs = CLng((((AvailableSerfs * 2) / 2) / 2) / 2)
       ElseIf GL_LEVEL = "GL4" Then
           TLogs = CLng(((((AvailableSerfs * 2) / 2) / 2) / 2) / 2)
       ElseIf GL_LEVEL = "GL5" Then
           TLogs = CLng((((((AvailableSerfs * 2) / 2) / 2) / 2) / 2) / 2)
       End If
       OutLine = OutLine & " Logs - " & TLogs & ","
       Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LOG", "ADD", TLogs)
       If GL_LEVEL = "GL0" Then
           TBark = CLng(AvailableSerfs * 5)
       ElseIf GL_LEVEL = "GL1" Then
           TBark = CLng((AvailableSerfs * 5) / 2)
       ElseIf GL_LEVEL = "GL2" Then
           TBark = CLng(((AvailableSerfs * 5) / 2) / 2)
       ElseIf GL_LEVEL = "GL3" Then
           TBark = CLng((((AvailableSerfs * 5) / 2) / 2) / 2)
       ElseIf GL_LEVEL = "GL4" Then
           TBark = CLng(((((AvailableSerfs * 5) / 2) / 2) / 2) / 2)
       ElseIf GL_LEVEL = "GL5" Then
           TBark = CLng((((((AvailableSerfs * 5) / 2) / 2) / 2) / 2) / 2)
       End If
       OutLine = OutLine & " Bark - " & TBark & ","
       Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "BARK", "ADD", TBark)
   End If
        
   ' Get tribute (silver)
   If GL_LEVEL = "GL0" Then
       Silver_Tribute = CLng(AvailableSerfs * 5)
   ElseIf GL_LEVEL = "GL1" Then
       Silver_Tribute = CLng((AvailableSerfs * 5) / 2)
   ElseIf GL_LEVEL = "GL2" Then
       Silver_Tribute = CLng(((AvailableSerfs * 5) / 2) / 2)
   ElseIf GL_LEVEL = "GL3" Then
       Silver_Tribute = CLng((((AvailableSerfs * 5) / 2) / 2) / 2)
   ElseIf GL_LEVEL = "GL4" Then
       Silver_Tribute = CLng(((((AvailableSerfs * 5) / 2) / 2) / 2) / 2)
   ElseIf GL_LEVEL = "GL5" Then
       Silver_Tribute = CLng((((((AvailableSerfs * 5) / 2) / 2) / 2) / 2) / 2)
   End If
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SILVER", "ADD", Silver_Tribute)
   OutLine = OutLine & " Silver - " & Silver_Tribute & ","
       
   TurnActOutPut = TurnActOutPut & OutLine
   
End If

ERR_Perform_Politics_Activities_CLOSE:
   Exit Function

ERR_Perform_Politics_Activities:
   Call A999_ERROR_HANDLING
   Resume ERR_Perform_Politics_Activities_CLOSE

End Function

Public Function Write_Book()
On Error GoTo ERR_Write_Book
TRIBE_STATUS = "Write_Book"

Dim DLLEVEL As Integer
Dim PARCHMENT As Long
Dim CHANCE As Integer
Dim HEXMAP As String

Set RESEARCHTABLE = TVDB.OpenRecordset("RESEARCH")
RESEARCHTABLE.index = "TOPIC"
RESEARCHTABLE.MoveFirst

Set TRIBESBOOKS = TVDBGM.OpenRecordset("TRIBES_BOOKS")
TRIBESBOOKS.index = "PRIMARYKEY"

BOOKTOPIC = Forms![WRITE BOOK]![TOPIC]
CLANNUMBER = Forms![WRITE BOOK]![CLAN]
TTRIBENUMBER = Forms![WRITE BOOK]![TRIBE]
HEXMAP = Forms![WRITE BOOK]![HEXMAP]

'DoCmd.Minimize

RESEARCHTABLE.MoveFirst
RESEARCHTABLE.Seek "=", BOOKTOPIC

RESEARCH_FOUND = "N"
Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, BOOKTOPIC)
        
If RESEARCH_FOUND = "Y" Then
   DLLEVEL = RESEARCHTABLE![DL REQUIRED]
   PARCHMENT = DLLEVEL * 10
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PARCHMENT", "SUBTRACT", PARCHMENT)
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "LEATHER", "SUBTRACT", 1)
   SKILLSTABLE.MoveFirst
   SKILLSTABLE.Seek "=", Skill_Tribe, TActivity
    
   If Not SKILLSTABLE.NoMatch Then
      CHANCE = SKILLSTABLE![SKILL LEVEL] * 5
   End If
   
   Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
   HEXMAPCONST.index = "FORTHKEY"
   If Not HEXMAPCONST.EOF Then
      HEXMAPCONST.MoveFirst
   End If
   HEXMAPCONST.Seek "=", HEXMAP, TCLANNUMBER, "LIBRARY"

   If Not HEXMAPCONST.NoMatch Then
      If HEXMAPCONST![1] > 0 Then
         CHANCE = CHANCE + 50
      End If
   End If
   
   DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 1, 0)

   If DICE1 <= CHANCE Then
      TRIBESBOOKS.Seek "=", TCLANNUMBER, TTRIBENUMBER, BOOKTOPIC
      If TRIBESBOOKS.NoMatch Then
         TRIBESBOOKS.AddNew
         TRIBESBOOKS![CLAN] = TCLANNUMBER
         TRIBESBOOKS![TRIBE] = TTRIBENUMBER
         TRIBESBOOKS![BOOK] = BOOKTOPIC
         TRIBESBOOKS![NUMBER] = 1
         TRIBESBOOKS.UPDATE
         TurnActOutPut = TurnActOutPut & " Book written, "
      Else
         TRIBESBOOKS.Edit
         TRIBESBOOKS![NUMBER] = TRIBESBOOKS![NUMBER] + 1
         TRIBESBOOKS.UPDATE
         TurnActOutPut = TurnActOutPut & " Book written, "
      End If
   Else
      TurnActOutPut = TurnActOutPut & " Book not written, "
   End If
Else
    Exit Function
End If

ERR_Write_Book_CLOSE:
   Exit Function

ERR_Write_Book:
   Call A999_ERROR_HANDLING
   Resume ERR_Write_Book_CLOSE

End Function


Public Function Perform_Armour_Making()
On Error GoTo ERR_Perform_Armour_Making
TRIBE_STATUS = "Perform_Armour_Making"

Initial_Armourers = TActives
TArmourers = TActives
     
 Call Process_Implement_Usage(TActivity, TItem, TArmourers, "NO")
       
 TribesSpecialists.MoveFirst
 TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ARMOURER"
       
 ' check availability of specialist.
 If Not TribesSpecialists.NoMatch Then
    TImplement = InputBox("How many Specialist Armourers Used?", "ARMOUR", "0")
    If TImplement > 0 Then
       If Not IsNull(TribesSpecialists![SPECIALISTS_USED]) Then
          total_available = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
       Else
          total_available = TribesSpecialists![SPECIALISTS]
       End If
       If TImplement > total_available Then
          TImplement = total_available
       End If
       If TImplement > TActives Then
          TActives = TActives + TActives
          TImplement = TImplement - TActives
          TribesSpecialists.Edit
          TribesSpecialists![SPECIALISTS_USED] = TribesSpecialists![SPECIALISTS_USED] + TActives
          TribesSpecialists.UPDATE
       Else
          TActives = TActives + TImplement
          TribesSpecialists.Edit
          TribesSpecialists![SPECIALISTS_USED] = TribesSpecialists![SPECIALISTS_USED] + TImplement
          TribesSpecialists.UPDATE
       End If
    End If
 End If
           
 Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
 
ERR_Perform_Armour_Making_CLOSE:
   Exit Function

ERR_Perform_Armour_Making:
   Call A999_ERROR_HANDLING
   Resume ERR_Perform_Armour_Making_CLOSE

End Function

Public Function CHECK_FOR_BUILDING(Building As String)
On Error GoTo ERR_CHECK_FOR_BUILDING
TRIBE_STATUS = "CHECK_FOR_BUILDING"

BUILDING_FOUND = "N"

Set VALID_CONST = TVDB.OpenRecordset("VALID_BUILDINGS")
VALID_CONST.index = "PRIMARYKEY"
VALID_CONST.MoveFirst
VALID_CONST.Seek "=", Building

'Tribes_Current_Hex is the current hex for the tribe
'Goods_Tribe_Current_Hex is the current hex for the goods tribe

HEXMAPCONST.index = "FORTHKEY"
HEXMAPCONST.Seek "=", Meeting_House_Hex, TCLANNUMBER, Building

If HEXMAPCONST.NoMatch Then
   BUILDING_FOUND = "N"
Else
   BUILDING_FOUND = "Y"
   If Building = "APIARY" Then
      APIARYS_FOUND = HEXMAPCONST![1]
   ElseIf Building = "HOSPITAL" Then
      HOSPITAL_FOUND = HEXMAPCONST![1]
   ElseIf Building = "MEETING HOUSE" Then
      MEETING_HOUSE_FOUND = HEXMAPCONST![1]
'   ElseIf BUILDING = "MILL" Then
'      MILL_FOUND = HEXMAPCONST![1]
   ElseIf Building = "SEWER" Then
      SEWER_FOUND = HEXMAPCONST![1]
   End If
   
' check which building is available for use and its capacity i.e. distillery
   TRIBE_STATUS = "CHECK_FOR_BUILDING_USAGE"
   Building_Used.index = "PRIMARYKEY"
   Building_Used.Seek "=", Meeting_House_Hex, TCLANNUMBER, Building
  
   If Building_Used.NoMatch Then
      Building_Used.AddNew
      Building_Used![Current_HexMap] = Meeting_House_Hex
      Building_Used![CLAN] = TCLANNUMBER
      Building_Used![Building] = Building
      Building_Used![USED] = 1
      Building_Used.UPDATE
      TBuildingLimit = HEXMAPCONST![1] * VALID_CONST![LIMITS]
   Else
      If Building_Used![USED] = 1 Then
         If HEXMAPCONST![2] > 0 Then
            TBuildingLimit = HEXMAPCONST![2] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 2 Then
         If HEXMAPCONST![3] > 0 Then
            TBuildingLimit = HEXMAPCONST![3] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 3 Then
         If HEXMAPCONST![4] > 0 Then
            TBuildingLimit = HEXMAPCONST![4] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 4 Then
         If HEXMAPCONST![5] > 0 Then
            TBuildingLimit = HEXMAPCONST![5] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 5 Then
         If HEXMAPCONST![6] > 0 Then
            TBuildingLimit = HEXMAPCONST![6] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 6 Then
         If HEXMAPCONST![7] > 0 Then
            TBuildingLimit = HEXMAPCONST![7] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 7 Then
         If HEXMAPCONST![8] > 0 Then
            TBuildingLimit = HEXMAPCONST![8] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 8 Then
         If HEXMAPCONST![9] > 0 Then
            TBuildingLimit = HEXMAPCONST![9] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] = 9 Then
         If HEXMAPCONST![10] > 0 Then
            TBuildingLimit = HEXMAPCONST![10] * VALID_CONST![LIMITS]
         End If
      ElseIf Building_Used![USED] >= 10 Then
         TBuildingLimit = 0
      End If
      Building_Used.Edit
      Building_Used![USED] = Building_Used![USED] + 1
      Building_Used.UPDATE
   End If
   ' HEXMAPCONST![N] may be -1 meaning building was not built. 0 means it is built but has no installations
   If HEXMAPCONST![1] > 0 Then TManufacturingLimit = HEXMAPCONST![1]
   If HEXMAPCONST![2] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![2]
   If HEXMAPCONST![3] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![3]
   If HEXMAPCONST![4] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![4]
   If HEXMAPCONST![5] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![5]
   If HEXMAPCONST![6] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![6]
   If HEXMAPCONST![7] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![7]
   If HEXMAPCONST![8] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![8]
   If HEXMAPCONST![9] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![9]
   If HEXMAPCONST![10] > 0 Then TManufacturingLimit = TManufacturingLimit + HEXMAPCONST![10]
   TManufacturingLimit = TManufacturingLimit * VALID_CONST![LIMITS]
End If

' check for limits
' how many ?? per building that is next to be used
' also, check for which building has been used








ERR_CHECK_FOR_BUILDING_CLOSE:
   VALID_CONST.Close
   Exit Function

ERR_CHECK_FOR_BUILDING:
   Call A999_ERROR_HANDLING
   Resume ERR_CHECK_FOR_BUILDING_CLOSE

End Function


Public Function A150_INITIALISE()
On Error GoTo ERR_A150_INITIALISE

TRIBE_STATUS = "A150_INITIALISE"
DebugOP "A150_INITIALISE"

Current_Turn = globalinfo![CURRENT TURN]
TURN_NUMBER = "TURN" & Left(globalinfo![CURRENT TURN], 2)

SLAVES_OVERSEEN = "N"
Firstmodify = "N"
LINENUMBER = 1
LINEFEED = Chr(13) & Chr(10)
TOTAL_FLAX = 0
TOTAL_GRAIN = 0
TOTAL_LINSEED = 0
TOTAL_COTTON = 0
TOTAL_SUGAR = 0
TOTAL_GRAPES = 0
TOTAL_TOBACCO = 0
TOTAL_HEMP = 0
TOTAL_POTATOES = 0
TOTAL_CROP = 0
PLANTING_STARTED = "NO"
ACRES_NOT_PLANTED = 0
ACRES_HARVESTED = 0
ACRES_TO_PLANT = 0

TFishing = 0
SEASON_FISHING = 0
GoatsKilled = 0
CattleKilled = 0
HorsesKilled = 0
DogsKilled = 0
CamelsKilled = 0
TNewHorses = 0

' Get Turn_Activity record if it exists
OutPutTable.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ACTIVITIES", 1

If OutPutTable.NoMatch Then
   TurnActOutPut = "^BTribe Activities:^B "
Else
   TurnActOutPut = OutPutTable![line detail]
End If

MORALELOSS = "N"
SKILL_SHORTAGE = 0
Number_of_Seeking_Groups = 0
Number_of_Seeking_Attempts = 0

TLogs = 0
TBark = 0
TDOGS = 0

Index1 = 1
Do Until Index1 > 4
   TSkillok(Index1) = "N"
   TSkill(Index1) = "N"
   TSkilllvl(Index1) = 0
   Index1 = Index1 + 1
Loop
    
Index1 = 1
Do Until Index1 > 40
   TGoods(Index1) = "empty"
   TQuantity(Index1) = 0
   Index1 = Index1 + 1
Loop
              
ModifyTable = "N"

ERR_A150_INITIALISE_CLOSE:
   Exit Function


ERR_A150_INITIALISE:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A150_INITIALISE_CLOSE
End If
End Function

Public Function A500_MAIN_PROCESS()
On Error GoTo ERR_A500_MAIN_PROCESS
TRIBE_STATUS = "A500_MAIN_PROCESS"
DebugOP "A500_MAIN_PROCESS"

ACTIVES_INUSE = TActives
             
Select Case TActivity
Case "Apiarism"
   TRIBE_STATUS = "Perform Apiarism"
   Call PERFORM_APIARISM
     
Case "Apothecary"
   TRIBE_STATUS = "Perform Apothecary"
   Call PERFORM_APOTHECARY
Case "ARMOUR"
   TRIBE_STATUS = "Perform Armour Making"
   If TSkillok(1) = "Y" Then
      Perform_Armour_Making
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Armour Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Armour Skill", "", 0, "NO")
      End If
   End If
     
Case "ART"
   TRIBE_STATUS = "Perform Art"
   Call Check_Available_Actives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", ", Art Activity Performed", "", 0, "YES")
   Else
      Call Check_Turn_Output(",", ", Art Activity Performed", "", 0, "YES")
   End If
   
Case "ATHEISM"
   TRIBE_STATUS = "Perform Atheism"
   ACTIVES_INUSE = TActives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", ", Festival Held", "", 0, "NO")
   Else
      Call Check_Turn_Output(",", ", Festival Held", "", 0, "NO")
   End If
   Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
   Call A150_Open_Tables("TRIBES_GENERAL_INFO")
     
Case "BAKING"
   TRIBE_STATUS = "Perform Baking"
   If TSkillok(1) = "Y" Then
      Call PERFORM_BAKING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Baking Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Baking Skill", "", 0, "NO")
      End If
   End If
    
Case "BAMBOOWORK"
   TRIBE_STATUS = "Perform Bamboowork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_COMMON("N", "Y", "Y", 3, "NONE")
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Bamboowork Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Bamboowork Skill", "", 0, "NO")
      End If
   End If
    
Case "Blubberwork"
   TRIBE_STATUS = "Perform Blubberwork"
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If
   ACTIVES_INUSE = TActives
   Call PERFORM_BLUBBERWORK

Case "BONEWORK"
   TRIBE_STATUS = "Perform Bonework"
   If TSkillok(1) = "Y" Then
      Call PERFORM_BONEWORK
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Bonework Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Bonework Skill", "", 0, "NO")
      End If
   End If
    
Case "BONING"
   TRIBE_STATUS = "Perform Boning"
   If TSkillok(1) = "Y" Then
      Call PERFORM_BONING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Boning Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Boning Skill", "", 0, "NO")
      End If
   End If

Case "BRICK MAKING"
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "BRICKLAYER")

   If NO_SPECIALISTS_FOUND > 0 Then
      If NO_SPECIALISTS_FOUND > TSpecialists Then
         NO_SPECIALISTS_FOUND = TSpecialists
      ElseIf NO_SPECIALISTS_FOUND < TSpecialists Then
         TSpecialists = NO_SPECIALISTS_FOUND
      End If
     
      Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "BRICKLAYER", "SPECIALISTS_USED", TSpecialists)
         
   End If
   
   TActives = TActives + TSpecialists
   
   TRIBE_STATUS = "Perform Brick Making"
   If TSkillok(1) = "Y" Then
      Call PERFORM_COMMON("Y", "Y", "Y", 3, "NONE")
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Brickmaking Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Brickmaking Skill", "", 0, "NO")
      End If
   End If

Case "CHEESE MAKING"
   TRIBE_STATUS = "Perform Cheese Making"
   If TSkillok(1) = "Y" Then
      Call PERFORM_CHEESE_MAKING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Cheesemaking Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Cheesemaking Skill", "", 0, "NO")
      End If
   End If

Case "CONVERT"
   TRIBE_STATUS = "Perform Conversions"
   Call Perform_Conversions
   
Case "COOKING"
   TRIBE_STATUS = "Perform Cooking"
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If

   Call Perform_Cooking
   Call Check_Available_Actives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", " Cooking Activity Performed", "", 0, "NO")
   Else
      Call Check_Turn_Output(",", " Cooking Activity Performed", "", 0, "NO")
   End If
   
Case "CURING"
   TRIBE_STATUS = "Perform Curing"
   'check maximum actives allocated and no more
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If
         
   If TSkillok(1) = "Y" Then
      Call PERFORM_CURING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Curing Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Curing Skill", "", 0, "NO")
      End If
   End If

Case "DANCING"
   TRIBE_STATUS = "Perform Dancing"
   Call Check_Available_Actives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", " Dancing Activity Performed", "", 0, "NO")
   Else
      Call Check_Turn_Output(",", " Dancing Activity Performed", "", 0, "NO")
   End If
   Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
   Call A150_Open_Tables("TRIBES_GENERAL_INFO")
 
Case "DEFAULT"
   TRIBE_STATUS = "Perform Default"
   Call SETUP_DEFAULT
     
Case "Defence"
   TRIBE_STATUS = "Perform Defence"
   ' NEED TO IDENTIFY WHAT ITEMS ARE BEING USED
     
    
Case "DISTILLING"
   TRIBE_STATUS = "Perform Distilling"
   If TSkillok(1) = "Y" Then
      Call PERFORM_DISTILLING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Distilling Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Distilling Skill", "", 0, "NO")
      End If
   End If

Case "DRESSING"
   TRIBE_STATUS = "Perform Dressing"
   If TSkillok(1) = "Y" Then
      Call PERFORM_DRESSING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Dressing Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Dressing Skill", "", 0, "NO")
      End If
   End If

Case "DRINKING"
   TRIBE_STATUS = "Perform Drinking"
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TItem, "SUBTRACT", TActives)
   Call Check_Available_Actives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", " Drinking Activity Performed", "", 0, "NO")
   Else
      Call Check_Turn_Output(",", " Drinking Activity Performed", "", 0, "NO")
   End If
   Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
   Call A150_Open_Tables("TRIBES_GENERAL_INFO")
  
Case "EATING"
   TRIBE_STATUS = "Perform Eating"
   Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TItem, "SUBTRACT", TActives)
   Call Check_Available_Actives
   If Right(TurnActOutPut, 3) = "^B " Then
      Call Check_Turn_Output("", " Eating Activity Performed", "", 0, "NO")
   Else
      Call Check_Turn_Output(",", " Eating Activity Performed", "", 0, "NO")
   End If
   If TActives >= (TMouths / 10) Then
       Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
       Call A150_Open_Tables("TRIBES_GENERAL_INFO")
   End If

Case "Engineering"
   TRIBE_STATUS = "Perform Engineering"
   TACLAN = TCLANNUMBER
   TAACTIVITY = TActivity
   TAITEM = TItem
   TADISTINCTION = TDistinction
 ' Joint projects?
    sCheckResult = isCheckBuildingEligibility(CONST_Tribes_Current_Hex, TOwning_Clan, TOwning_Tribe, TItem)
    If sCheckResult = "True" Then
        Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "ENGINEER")
   
        If TSpecialists > NO_SPECIALISTS_FOUND Then
            TSpecialists = NO_SPECIALISTS_FOUND
        End If
  
        Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "ENGINEER", "SPECIALISTS_USED", TSpecialists)

   ' Actives & Specialists are combined prior to entering common function
        TActives = TActives + TSpecialists
   
        If TJoint = "Y" Then
            TATRIBE = TOwning_Tribe
            TActives = (TActives * (10 / (10 + SKILL_SHORTAGE)))
        ElseIf SKILL_SHORTAGE > 0 Then
            If Right(TurnActOutPut, 3) = "^B " Then
                Call Check_Turn_Output("", " Insufficient skill level for Engineering", "", 0, "NO")
            Else
                Call Check_Turn_Output(",", " Insufficient skill level for Engineering", "", 0, "NO")
            End If
        Else
            TATRIBE = TTRIBENUMBER
        End If
  
        If TItem = "MOAT" Or TItem = "DITCH" Then
            ImplementsTable.index = "ACTIVITY"
            ImplementsTable.MoveFirst
            ImplementsTable.Seek "=", "ENGINEERING", "ALL"
            IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
      
            Do While ImplementsTable![ACTIVITY] = TActivity
                QUESTION = "How many " & ImplementsTable![IMPLEMENT]
                QUESTION = QUESTION & " used for building " & TAITEM & "?"
                Call Process_Implement_Usage(TAACTIVITY, ImplementsTable![IMPLEMENT], TActives, "YES")
                ImplementsTable.MoveNext
                If ImplementsTable.EOF Then
                    Exit Do
                End If
                IMPLEMENT_MODIFIER = ImplementsTable![Modifier]
            Loop
      
            ImplementsTable.index = "PRIMARYKEY"
            ImplementsTable.MoveFirst
        Else
            Call Process_Implement_Usage(TActivity, "ALL", TActives, "NO")
        End If
     
   ' Allow for specialists double benefit
        TActives = TActives + TSpecialists
   
        Call Calc_Engineering(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TActives)
    Else
        Call Check_Turn_Output(",", sCheckResult, "", 0, "NO")
        Msg = "The Clan was " & TCLANNUMBER & " The Tribe was " & TTRIBENUMBER
        Msg = Msg & Chr(13) & Chr(10) & sCheckResult
        MsgBox (Msg)
    End If

Case "EXCAVATION"
   TRIBE_STATUS = "Perform Excavation"
   If TSkillok(1) = "Y" Then
      TurnActOutPut = TurnActOutPut & ", Excavate "
      Call Perform_Excavation
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Excavation", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Excavation", "", 0, "NO")
      End If
   End If

Case "Fishing"
   TRIBE_STATUS = "Perform Fishing"
   Call PERFORM_FISHING
    
Case "FLENSING"
   TRIBE_STATUS = "Perform Flensing"
   If TSkillok(1) = "Y" Then
      Call PERFORM_FLENSING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Flensing Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Flensing Skill", "", 0, "NO")
      End If
   End If

Case "FLENSING&PEELING"
   TRIBE_STATUS = "Perform Flensing & Peeling"
   Call SET_SKILL_LEVEL_1(FLENSING_LEVEL)
   Call SET_SKILL_LEVEL_2(PEELING_LEVEL)
   Call PERFORM_FLENSING_AND_PEELING

Case "FLENSING&PEELING&BONING"
   TRIBE_STATUS = "Perform Flensing & Peeling & Boning"
   Call SET_SKILL_LEVEL_1(FLENSING_LEVEL)
   Call SET_SKILL_LEVEL_2(PEELING_LEVEL)
   Call SET_SKILL_LEVEL_3(BONING_LEVEL)
   Call PERFORM_FLENSING_AND_PEELING_AND_BONING

Case "FLETCHING"
   TRIBE_STATUS = "Perform Flensing"
   If TSkillok(1) = "Y" Then
      Call PERFORM_FLETCHING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Fletching Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Fletching Skill", "", 0, "NO")
      End If
   End If

Case "Foraging"
   TRIBE_STATUS = "Perform Foraging"
   ACTIVES_INUSE = TActives
   Call PERFORM_GATHERING
    
Case "FORESTRY"
   TRIBE_STATUS = "Perform Forestry"
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If
   ACTIVES_INUSE = TActives
   If TItem = "CHARCOAL MAKING" Then
      Call PERFORM_FORESTRY
   Else
      Select Case TRIBES_TERRAIN
      Case "CONIFER HILLS"
         Call PERFORM_FORESTRY
      Case "DECIDUOUS"
         Call PERFORM_FORESTRY
      Case "DECIDUOUS FLAT"
         Call PERFORM_FORESTRY
      Case "DECIDUOUS FOREST"
         Call PERFORM_FORESTRY
      Case "DECIDUOUS HILLS"
         Call PERFORM_FORESTRY
      Case "HARDWOOD FOREST"
         Call PERFORM_FORESTRY
      Case "JUNGLE"
         Call PERFORM_FORESTRY
      Case "JUNGLE HILLS"
         Call PERFORM_FORESTRY
      Case "LOW CONIFER MOUNTAINS"
         Call PERFORM_FORESTRY
      Case "LOW CONIFER MT"
         Call PERFORM_FORESTRY
      Case "LOW JUNGLE MOUNTAINS"
         Call PERFORM_FORESTRY
      Case "LOW JUNGLE MT"
         Call PERFORM_FORESTRY
      Case "MANGROVE SWAMP"
         Call PERFORM_FORESTRY
      Case Else
        If Right(TurnActOutPut, 3) = "^B " Then
           Call Check_Turn_Output("", " Invalid terrain for Forestry", "", 0, "NO")
        Else
           Call Check_Turn_Output(",", " Invalid terrain for Forestry", "", 0, "NO")
        End If
      End Select
   End If

Case "Furrier"
   TRIBE_STATUS = "Perform Furrier"
   Call PERFORM_FURRIER

Case "Gathering"
   TRIBE_STATUS = "Perform Gathering"
   ACTIVES_INUSE = TActives
   If TItem = "CLAY" Then
      Call PERFORM_POTTERY
   Else
      Call PERFORM_GATHERING
   End If
       
Case "GLASSWORK"
   TRIBE_STATUS = "Perform Glasswork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_GLASSWORK
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Glasswork Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Glasswork Skill", "", 0, "NO")
      End If
   End If

Case "Gut&Bone"
   TRIBE_STATUS = "Perform Gut & Bone"
   Call SET_SKILL_LEVEL_1(GUTTING_LEVEL)
   Call SET_SKILL_LEVEL_2(BONING_LEVEL)
       
   If MAXIMUM_ACTIVES_1 >= (TActives / 2) Then
      If MAXIMUM_ACTIVES_2 >= (TActives / 2) Then
         ACTIVES_INUSE = TActives
         Call PERFORM_GUT_AND_BONE
      Else
         TActives = MAXIMUM_ACTIVES_2
         ACTIVES_INUSE = TActives
         Call PERFORM_GUT_AND_BONE
      End If
   Else
      TActives = MAXIMUM_ACTIVES_1
      ACTIVES_INUSE = TActives
      Call PERFORM_GUT_AND_BONE
   End If
    
Case "GUT&SKIN"
   TRIBE_STATUS = "Perform Gut & Skin"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(GUTTING_LEVEL)
   Call PERFORM_SKIN_AND_GUT
  
Case "Gutting"
   TRIBE_STATUS = "Perform Gutting"
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If
   ACTIVES_INUSE = TActives
   Call PERFORM_GUTTING
    
Case "HARVEST"
   TRIBE_STATUS = "Perform Harvest"
   ACRES_PLANTED = 0
   ACRES_HARVESTED = 0
   Call PERFORM_HARVESTING(CURRENT_CLIMATE, TItem)

Case "Healing"
   TRIBE_STATUS = "Perform Healing"
   Call PERFORM_HEALING
       
Case "Herding"
   TRIBE_STATUS = "Perform Herding"
   Call PERFORM_HERDING
    
Case "Hunting"
   TRIBE_STATUS = "Perform Hunting"
   Call PERFORM_HUNTING
       
Case "JEWELLERY"
   TRIBE_STATUS = "Perform Jewellery"
   If TSkillok(1) = "Y" Then
      Call PERFORM_JEWELLERY
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Jewellery Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Jewellery Skill", "", 0, "NO")
      End If
   End If

Case "KILLING"
   TRIBE_STATUS = "Perform Killing"
   Call PERFORM_KILLING

Case "LEATHERWORK"
   TRIBE_STATUS = "Perform Leatherwork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_LEATHERWORK
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Leatherwork Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Leatherwork Skill", "", 0, "NO")
      End If
   End If

Case "LITERACY"
   TRIBE_STATUS = "Perform Literacy"
   Call PERFORM_BOOK_WRITING
       
Case "MAINTAINING"
   TRIBE_STATUS = "Perform Maintaining"
 
   If ((Left(Current_Turn, 2) >= 7) And (Left(Current_Turn, 2) <= 12)) Then
      TurnActOutPut = TurnActOutPut & " wrong month to be maintaining " & TItem
      GoTo EXIT_MAINTAIN
   End If
   
   ' Identify the Crop Status.
   Set CROP_TABLE = TVDB.OpenRecordset("VALID_CROPS")
   CROP_TABLE.index = "PRIMARYKEY"
   CROP_TABLE.MoveFirst
   CROP_TABLE.Seek "=", TItem
 
   If CROP_TABLE.NoMatch Then
      MsgBox (TItem & " not on Valid_Crops table")
      ACRES_TO_MAINTAIN = 1
   Else
      ACRES_TO_MAINTAIN = CROP_TABLE![ACRES_MAINTAINED]
   End If

   CROP_TABLE.Close
       
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FARMER")
   
   If TSpecialists > FARMER_FOUND Then
      TSpecialists = FARMER_FOUND
   End If
  
   Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "FARMER", "SPECIALISTS_USED", TSpecialists)

   ' Actives & Specialists are combined prior to entering common function
   TActives = TActives + (TSpecialists * 2)
   
    PermFarmingTable.MoveFirst
    PermFarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, TItem
     
    Do While TActives > 0
         ACRES_MAINTAINED = ACRES_MAINTAINED + ACRES_TO_MAINTAIN
         TActives = TActives - 1
    Loop
       
    TurnActOutPut = TurnActOutPut & ", " & ACTIVES_INUSE & " Maintained " & ACRES_MAINTAINED
    If TItem = "HERBS" Or TItem = "HASHISH" Then
       TurnActOutPut = TurnActOutPut & " plots of " & TItem
    Else
       TurnActOutPut = TurnActOutPut & " acres " & TItem
    End If

    'if the crop is Hashish and it is turn 3, then harvest 1000
    If TItem = "HASHISH" Then
       If Left(Current_Turn, 2) = "03" Then
          Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TItem, "ADD", 1000)
       End If
    ElseIf TItem = "GRAPES" Then
       ' ??
       'Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, Goods_Tribe, TItem, "ADD", 1000)
    ElseIf TItem = "HERBS" Then
       ' ??
       'Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, Goods_Tribe, TItem, "ADD", 1000)
    End If
    
EXIT_MAINTAIN:

Case "METALWORK"
   TRIBE_STATUS = "Perform Metalwork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_METALWORK
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Metalwork Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Metalwork Skill", "", 0, "NO")
      End If
   End If

Case "MILLING"
   TRIBE_STATUS = "Perform Milling"
   If TSkillok(1) = "Y" Then
      Call PERFORM_MILLING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Milling Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Milling Skill", "", 0, "NO")
      End If
   End If

Case "Mining"
   TRIBE_STATUS = "Perform Mining"
   Call PERFORM_MINING

Case "MUSIC"
   TRIBE_STATUS = "Perform Music"
   If TItem = "PERFORM" Then
      If TActives >= (TMouths / 10) Then
         Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
         Call A150_Open_Tables("TRIBES_GENERAL_INFO")
      End If
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " Music Activity Performed", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " Music Activity Performed", "", 0, "NO")
      End If
   ElseIf TSkillok(1) = "Y" Then
      Call PERFORM_MUSIC
   End If

Case "PACIFICATION"
   TRIBE_STATUS = "Perform Pacification"
   ' PERFORMED IN FINAL ACTIVITIES
   ' JUST CATERED FOR HERE TO IDENTIFY WARRIOR USAGE.

Case "PEELING"
   TRIBE_STATUS = "Perform Peeling"
   If TSkillok(1) = "Y" Then
      Call PERFORM_PEELING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Peeling Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Peeling Skill", "", 0, "NO")
      End If
   End If

 Case "Planting"
   TRIBE_STATUS = "Perform Planting"
   ' Identify the Crop Status.
   Set CROP_TABLE = TVDB.OpenRecordset("VALID_CROPS")
   CROP_TABLE.index = "PRIMARYKEY"
   CROP_TABLE.MoveFirst
   CROP_TABLE.Seek "=", TItem
 
   ACRES_PLANTED = 0
   ACRES_TO_PLANT = 0
   
   If CROP_TABLE.NoMatch Then
      MsgBox (TItem & " not on Valid_Crops table")
      ACRES_TO_PLANT = 1
   Else
      ACRES_TO_PLANT = CROP_TABLE![ACRES_PLANTING]
   End If

   CROP_TYPE = CROP_TABLE![CROP_TYPE]

   CROP_TABLE.Close
   
   ' get the number of specialists
   
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FARMER")
   
   If TSpecialists > FARMER_FOUND Then
      TSpecialists = FARMER_FOUND
   End If
  
   Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "FARMER", "SPECIALISTS_USED", TSpecialists)

   ' Actives & Specialists are combined prior to entering common function
   TActives = TActives + TSpecialists
   ACTIVES_INUSE = TActives
   
   Call Process_Implement_Usage("PLANTING", TItem, TActives, "YES")
   
    ' Allow for specialists double benefit
    TActives = TActives + TSpecialists
   
  If CROP_TYPE = "Temporary" Then
       ACRES_PLANTED = 0
       ' GET ACRES AVAILABLE
       FarmingTable.MoveFirst
       FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, Current_Turn, "PLOWED"
       ACRES_PLOWED = FarmingTable![ITEM_NUMBER]
          
       FarmingTable.MoveFirst
       FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, Current_Turn, TItem
       ' 3 acres
       If FarmingTable.NoMatch Then
           FarmingTable.AddNew
           FarmingTable![HEXMAP] = Tribes_Current_Hex
           FarmingTable![CLAN] = TCLANNUMBER
           FarmingTable![TRIBE] = TTRIBENUMBER
           FarmingTable![TURN] = Current_Turn
           FarmingTable![ITEM] = TItem
           FarmingTable![ITEM_NUMBER] = 0
           FarmingTable.UPDATE
       End If
       FarmingTable.MoveFirst
       FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, Current_Turn, TItem
       
       If PLANTING_STARTED = "NO" Then
           PLANTING_STARTED = "YES"
           ACRES_NOT_PLANTED = ACRES_PLOWED
       End If
            
       Do While TActives > 0
          FarmingTable.Edit
          If ACRES_NOT_PLANTED >= ACRES_TO_PLANT Then
              FarmingTable![ITEM_NUMBER] = FarmingTable![ITEM_NUMBER] + ACRES_TO_PLANT
              ACRES_PLANTED = ACRES_PLANTED + ACRES_TO_PLANT
              ACRES_NOT_PLANTED = ACRES_NOT_PLANTED - ACRES_TO_PLANT
          Else
              FarmingTable![ITEM_NUMBER] = FarmingTable![ITEM_NUMBER] + ACRES_NOT_PLANTED
              ACRES_PLANTED = ACRES_PLANTED + ACRES_NOT_PLANTED
              TActives = 0
          End If
          FarmingTable.UPDATE
          TActives = TActives - 1
      Loop
       
      TurnActOutPut = TurnActOutPut & ", " & ACTIVES_INUSE & " Planted " & ACRES_PLANTED
      TurnActOutPut = TurnActOutPut & " acres " & TItem
      
      FarmingTable.MoveFirst
      FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, Current_Turn, "PLOWED"
      ACRES_PLOWED = FarmingTable![ITEM_NUMBER]
      FarmingTable.Edit
      FarmingTable![ITEM_NUMBER] = FarmingTable![ITEM_NUMBER] - ACRES_PLANTED
      FarmingTable.UPDATE
   Else
      If TItem = "HERBS" Then
          Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "HERB")
          If TActives > Num_Goods Then
              TActives = Num_Goods
          End If
      End If
        
      If PermFarmingTable.BOF Then
          ' do nothing
      Else
           PermFarmingTable.MoveFirst
      End If
      PermFarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, TItem
       
      If PermFarmingTable.NoMatch Then
         PermFarmingTable.AddNew
         PermFarmingTable![HEXMAP] = Tribes_Current_Hex
         PermFarmingTable![CLAN] = TCLANNUMBER
         PermFarmingTable![TRIBE] = TTRIBENUMBER
         PermFarmingTable![ITEM] = TItem
         PermFarmingTable![ITEM_NUMBER] = 0
         PermFarmingTable.UPDATE
      End If
      PermFarmingTable.MoveFirst
      PermFarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, TItem
       
      Do While TActives > 0
          PermFarmingTable.Edit
          PermFarmingTable![ITEM_NUMBER] = PermFarmingTable![ITEM_NUMBER] + ACRES_TO_PLANT
          ACRES_PLANTED = ACRES_PLANTED + ACRES_TO_PLANT
          PermFarmingTable.UPDATE
          TActives = TActives - 1
      Loop
       
      TurnActOutPut = TurnActOutPut & ", " & ACTIVES_INUSE & " Planted " & ACRES_PLANTED
      TurnActOutPut = TurnActOutPut & " acres " & TItem
      
   End If

Case "Plowing"
   If TRIBES_TERRAIN = "PRAIRIE" Or TRIBES_TERRAIN = "GRASSY HILLS" Then
    If CheckFarmingEligibility(Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER) = "TRUE" Then
   
        TRIBE_STATUS = "Perform Plowing"
        ACRES_PLOWED = 0
      
        ' get the number of specialists
  
        Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "FARMER")
   
        If TSpecialists > FARMER_FOUND Then
            TSpecialists = FARMER_FOUND
        End If
  
        Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "FARMER", "SPECIALISTS_USED", TSpecialists)

        ' Actives & Specialists are combined prior to entering common function
        TActives = TActives + (TSpecialists * 2)

        Call Process_Implement_Usage("PLOWING", "ALL", ACRES_PLOWED, "YES")
        ImplementsTable.index = "PRIMARYKEY"
        ImplementsTable.MoveFirst
        If FarmingTable.BOF Then
        ' do nothing
        Else
            FarmingTable.MoveFirst
        End If
        FarmingTable.Seek "=", Tribes_Current_Hex, TCLANNUMBER, TTRIBENUMBER, Current_Turn, "PLOWED"
        If FarmingTable.NoMatch Then
            FarmingTable.AddNew
            FarmingTable![HEXMAP] = Tribes_Current_Hex
            FarmingTable![CLAN] = TCLANNUMBER
            FarmingTable![TRIBE] = TTRIBENUMBER
            FarmingTable![TURN] = Current_Turn
            FarmingTable![ITEM] = "PLOWED"
            FarmingTable![ITEM_NUMBER] = ACRES_PLOWED
            FarmingTable.UPDATE
        Else
            FarmingTable.Edit
            FarmingTable![ITEM_NUMBER] = ACRES_PLOWED
            FarmingTable.UPDATE
        End If
     
        If Right(TurnActOutPut, 3) = "^B " Then
            Call Check_Turn_Output(" Plowed ", " acres ", "", ACRES_PLOWED, "NO")
        Else
            Call Check_Turn_Output(", Plowed ", " acres ", "", ACRES_PLOWED, "NO")
   
        End If
    Else
        Call Check_Turn_Output(", Can't plow without nearby village ", " ", "", 0, "NO")
    End If
   Else
    Call Check_Turn_Output(", Can't plow ", TRIBES_TERRAIN, "", 0, "NO")
        Msg = Msg & "The Clan was " & TCLANNUMBER & " The Tribe was " & TTRIBENUMBER
        Msg = Msg & Chr(13) & Chr(10) & " This location is not suitable for plowing" 'CheckFarmingEligibility
        MsgBox (Msg)
   End If
Case "POLITICS"
   TRIBE_STATUS = "Perform Pacification"
   ' PERFORMED IN FINAL ACTIVITIES
   ' JUST CATERED FOR HERE TO IDENTIFY WARRIOR USAGE.

Case "POTTERY"
   TRIBE_STATUS = "Perform Pottery"
   If TSkillok(1) = "Y" Then
      Call PERFORM_POTTERY
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Pottery Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Pottery Skill", "", 0, "NO")
      End If
   End If

Case "Quarrying"
   TRIBE_STATUS = "Perform Quarrying"
   Call Get_Specialists_Info(TCLANNUMBER, TTRIBENUMBER, "QUARRIER")

   If TSpecialists > NO_SPECIALISTS_FOUND Then
      TSpecialists = NO_SPECIALISTS_FOUND
   End If
     
   Call UPDATE_TRIBES_SPECIALISTS(TCLANNUMBER, TTRIBENUMBER, "QUARRIER", "SPECIALISTS_USED", TSpecialists)
         
   TActives = TActives + TSpecialists
   
   QUARRYING_ACTIVES = TActives
   Select Case TRIBES_TERRAIN
   Case "PRAIRIE"
     If Right(TurnActOutPut, 3) = "^B " Then
        Call Check_Turn_Output("", " Invalid terrain for Quarrying", "", 0, "NO")
     Else
        Call Check_Turn_Output(",", " Invalid terrain for Quarrying", "", 0, "NO")
     End If
   Case Else
      If MAXIMUM_ACTIVES_1 < TActives Then
         TActives = MAXIMUM_ACTIVES_1
      End If
      ACTIVES_INUSE = TActives
      QUARRYING_ACTIVES = TActives
     
      If TItem = "MARBLE" Then
         If QUARRYING = "Y" Then
            Call Process_Implement_Usage("Quarrying", "MARBLE", QUARRYING_ACTIVES, "NO")
         End If
      ElseIf TItem = "SLATE" Then
          Call Process_Implement_Usage("Quarrying", "SLATE", QUARRYING_ACTIVES, "NO")
      Else
          Call Process_Implement_Usage("Quarrying", "ALL", QUARRYING_ACTIVES, "NO")
          If QUARRYING = "Y" Then
              QUARRYING_ACTIVES = QUARRYING_ACTIVES + QUARRYING_ACTIVES
          End If
      End If
                 
      ' allow for specialists double benefit
      QUARRYING_ACTIVES = QUARRYING_ACTIVES + TSpecialists
      
      ImplementsTable.index = "PRIMARYKEY"
      ImplementsTable.MoveFirst
        
      If TItem = "STONE" Then
         STONES_QUARRIED = (QUARRYING_ACTIVES * STONES_TO_QUARRY)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "STONE", "ADD", STONES_QUARRIED)
         Call Check_Turn_Output(",", " effective people dug ", " stone ", STONES_QUARRIED, "YES")
      ElseIf TItem = "FLAGSTONE" Then
         STONES_QUARRIED = (QUARRYING_ACTIVES * STONES_TO_QUARRY) * 2
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "FLAGSTONE", "ADD", STONES_QUARRIED)
         Call Check_Turn_Output(",", " effective people dug ", " flagstone ", STONES_QUARRIED, "YES")
      ElseIf TItem = "GRAVEL" Then
         STONES_QUARRIED = (QUARRYING_ACTIVES * 20)
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "GRAVEL", "ADD", STONES_QUARRIED)
         Call Check_Turn_Output(",", " effective people dug ", " gravel ", STONES_QUARRIED, "YES")
      ElseIf TItem = "MARBLE" Then
         If QUARRYING = "Y" Then
            QUARRYING_ACTIVES = QUARRYING_ACTIVES - ACTIVES_INUSE
            STONES_QUARRIED = QUARRYING_ACTIVES * 2
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "MARBLE", "ADD", STONES_QUARRIED)
            Call Check_Turn_Output(",", " effective people dug ", " marble ", STONES_QUARRIED, "YES")
         End If
      ElseIf TItem = "SLATE" Then
         QUARRYING_ACTIVES = QUARRYING_ACTIVES - ACTIVES_INUSE
         STONES_QUARRIED = QUARRYING_ACTIVES * 2
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SLATE", "ADD", STONES_QUARRIED)
         Call Check_Turn_Output(",", " effective people dug ", " slate ", STONES_QUARRIED, "YES")
      End If
        
      DoCmd.Hourglass True
   End Select
    
Case "Refining"
   TRIBE_STATUS = "Perform Refining"
   If MAXIMUM_ACTIVES_1 < TActives Then
      TActives = MAXIMUM_ACTIVES_1
   End If
   ACTIVES_INUSE = TActives
   Call PERFORM_REFINING

Case "RELIGION"
   TRIBE_STATUS = "Perform Religion"
   ACTIVES_INUSE = TActives
   If TItem = "HEADS" Then
      Select Case WEATHER
      Case "WIND"
         FIND_CHANCE = (3 + SCOUTING_LEVEL) - 1
      Case "L-SNOW"
         FIND_CHANCE = (3 + SCOUTING_LEVEL) - 1
      Case "H-SNOW"
         FIND_CHANCE = (3 + SCOUTING_LEVEL) - 2
      Case "H-RAIN"
         FIND_CHANCE = (3 + SCOUTING_LEVEL) - 2
      Case "L-RAIN"
         FIND_CHANCE = (3 + SCOUTING_LEVEL) - 2
      Case Else
         FIND_CHANCE = (3 + SCOUTING_LEVEL)
      End Select
      roll1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 1, 0)
      If roll1 <= FIND_CHANCE Then
         THeads = CLng(Sqr(TActives * ((RELIGION_LEVEL + 10) / 10)))
         THeads = CLng(THeads * ((COMBAT_LEVEL + 10) / 10))
         If THeads < 1 Then
            THeads = 1
         End If
         Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "HEADS", "ADD", THeads)
      End If
      DoCmd.Hourglass True
      ' Update output line
      If THeads > 0 Then
         Call Check_Turn_Output(",", " THeads ", "", THeads, "NO")
         THeads = 0
      Else
         Call Check_Turn_Output(",", " Find No Heads ", "", 0, "NO")
      End If
   Else
      Call Check_Turn_Output(",", " Festival Held ", "", 0, "NO")
      Call CHECK_MORALE(TCLANNUMBER, TTRIBENUMBER) ' FOUND IN GLOBAL FUNCTIONS MODULE
      Call A150_Open_Tables("TRIBES_GENERAL_INFO")
   End If
  
Case "RESEARCH"
   TRIBE_STATUS = "Perform Research"
   If TSkillok(1) = "Y" Then
      Call PERFORM_RESEARCH
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Research Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Research Skill", "", 0, "NO")
      End If
   End If

Case "Salting"
   TRIBE_STATUS = "Perform Salting"
   ' this needs to change
   ' fish are automatically salted - but is it working??
   ' this needs to cater for salting other items
   If TSkillok(1) = "Y" Then
      Call PERFORM_SALTING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Salting Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Salting Skill", "", 0, "NO")
      End If
   End If

   

Case "Sand Gathering"
   TRIBE_STATUS = "Perform Sand Gathering"
   ACTIVES_INUSE = TActives
   If TItem = "CLAY" Then
      Call PERFORM_POTTERY
   Else
      Call PERFORM_GATHERING
   End If
       
Case "SB"
   TRIBE_STATUS = "Perform SB"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(BONING_LEVEL)
   Call PERFORM_SKIN_AND_BONE
    
Case "SCOUTING"
   TRIBE_STATUS = "Perform Scouting"
   ' NEED TO IDENTIFY WHAT ITEMS ARE BEING USED

Case "SECURITY"
   TRIBE_STATUS = "Perform Security"
   ' NEED TO IDENTIFY WHAT ITEMS ARE BEING USED

Case "SEEKING"
   TRIBE_STATUS = "Perform Seeking"
   ' need to cater for a modifier for each item being found.  This will be devided into the calc to give the figure found.
   ' read seeking_returns_table for modifier
   ' need to record what a clan seeks and then ignore future attempts.
   ' need a table
   
   SeekingReturnsTable.MoveFirst
   SeekingReturnsTable.Seek "=", TItem
    
   Number_Of_Implements = 0
   
   If TActives > 100 Then
       TActives = 100
   End If
   
   Number_Found = TActives + (HORSES * 1.3)
   
   Call Process_Implement_Usage(TActivity, TItem, Number_Of_Implements, "NO")

   If Number_Of_Implements > 0 Then
      Number_Found = Number_Found + Number_Of_Implements
   End If
      
   Number_Found = Number_Found * (1 + SCOUTING_LEVEL / 3 + SEEKING_LEVEL / 2) / SeekingReturnsTable![Modifier]
   Number_Found = CLng(Number_Found / 12)
 
   If TItem = "RECRUITS" Then
      TRIBESINFO.Edit
      TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] + CLng(Number_Found)
      TRIBESINFO.UPDATE

   Else
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, TItem, "ADD", Number_Found)
   End If
   DoCmd.Hourglass True

   ' Update output line
   If Number_Found > 0 Then
       Msg = " " & TItem
       Call Check_Turn_Output(", Find ", Msg, "", Number_Found, "NO")
   Else
       Msg = ", Find No " & TItem
       TurnActOutPut = TurnActOutPut & Msg
   End If
       
Case "SEWING"
   TRIBE_STATUS = "Perform Sewing"
   If TSkillok(1) = "Y" Then
      Call PERFORM_SEWING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Sewing Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Sewing Skill", "", 0, "NO")
      End If
   End If

Case "SG"
   TRIBE_STATUS = "Perform SG"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(GUTTING_LEVEL)
   Call PERFORM_SKIN_AND_GUT
  
Case "SGB"
   TRIBE_STATUS = "Perform SGB"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(GUTTING_LEVEL)
   Call SET_SKILL_LEVEL_3(BONING_LEVEL)
   Call PERFORM_SKIN_AND_GUT_AND_BONE
    
Case "SHEARING"
   TRIBE_STATUS = "Perform Shearing"
   ' Get goats
   TNumItems = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "GOAT")
   ' actives * 10 = number of goats sheared
   ' Get actives = TActives
   If TActives * 10 >= TNumItems Then
      ' shear
   Else
      TNumItems = TActives * 10
   End If
   ' sheared goat * 15 = cotton
   Number_Found = TNumItems * 15
   If Left(Current_Turn, 2) = "06" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "COTTON", "ADD", Number_Found)
   ElseIf Left(Current_Turn, 2) = "12" Then
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "COTTON", "ADD", Number_Found)
   End If
     
Case "Skin&Gut"
   TRIBE_STATUS = "Perform Skin & Gut"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(GUTTING_LEVEL)
   Call PERFORM_SKIN_AND_GUT
  
Case "Skin&Bone"
   TRIBE_STATUS = "Perform Skin & Bone"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(BONING_LEVEL)
   Call PERFORM_SKIN_AND_BONE
    
Case "Skin&Gut&Bone"
   TRIBE_STATUS = "Perform Skin, Gut & Bone"
   Call SET_SKILL_LEVEL_1(SKINNING_LEVEL)
   Call SET_SKILL_LEVEL_2(GUTTING_LEVEL)
   Call SET_SKILL_LEVEL_3(BONING_LEVEL)
   Call PERFORM_SKIN_AND_GUT_AND_BONE
    
Case "shipbuilding"
   TRIBE_STATUS = "Perform Shipbuilding"
   TACLAN = TCLANNUMBER
   TAACTIVITY = TActivity
   TAITEM = TItem
   TADISTINCTION = TDistinction
   SHIPBUILDING_ACTIVES = TActives
   TempActives = TActives
   Call Process_Implement_Usage("Shipbuilding", "ALL", SHIPBUILDING_ACTIVES, "YES")
      
   ImplementsTable.index = "PRIMARYKEY"
   ImplementsTable.MoveFirst
      
   If TJoint = "Y" Then
      JOINT_TRIBE = InputBox("What Tribe is this Ship to belong to?", "TRIBE", "N")
      TATRIBE = JOINT_TRIBE
      TAACTIVES = (SHIPBUILDING_ACTIVES * (10 / (10 + SKILL_SHORTAGE)))
   Else
      TATRIBE = TTRIBENUMBER
      TAACTIVES = SHIPBUILDING_ACTIVES
   End If
    
   TActives = TempActives
    
   Call Calc_Shipbuilding(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TAACTIVES)
       
Case "SIEGE EQUIPMENT"
   TRIBE_STATUS = "Perform Siege Equipment"
   If TSkillok(1) = "Y" Then
      Call PERFORM_SIEGE_EQUIPMENT
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Siege Equipment Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Siege Equipment Skill", "", 0, "NO")
      End If
   End If

Case "SKINNING"
   TRIBE_STATUS = "Perform Skinning"
   If TSkillok(1) = "Y" Then
      Call PERFORM_SKINNING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Skinning Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Skinning Skill", "", 0, "NO")
      End If
   End If

Case "Slavery"
   TRIBE_STATUS = "Perform Slavery"
   PROCESSITEMS.MoveFirst
   PROCESSITEMS.Seek "=", TTRIBENUMBER, TActivity, TItem, "SHACKLE"
   If Not PROCESSITEMS.NoMatch Then
      ImplementUsage.MoveFirst
      ImplementUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, "SHACKLE"

      If Not ImplementUsage.NoMatch Then
          ' GET THE NUMBER OF SHACKLES ALLOCATED
          total_available = ImplementUsage![total_available] - ImplementUsage![Number_Used]
          If total_available > NUMBER_OF_SLAVES Then
             Call Update_Implement_Usage(TCLANNUMBER, GOODS_TRIBE, "SHACKLE", NUMBER_OF_SLAVES)
             NUMBER_OF_SLAVES = CLng(NUMBER_OF_SLAVES / 2)
          Else
             Call Update_Implement_Usage(TCLANNUMBER, GOODS_TRIBE, "SHACKLE", total_available)
             NUMBER_OF_SLAVES = NUMBER_OF_SLAVES - total_available
          End If
      End If
   End If
  
   SLAVES_OVERSEEN = "Y"
   'UPDATE TRIBE_PROCESSING WITH DETAILS.
   Tribes_Processing.Seek "=", TTRIBENUMBER
   If Tribes_Processing.NoMatch Then
      Tribes_Processing.AddNew
      Tribes_Processing![TRIBE] = TTRIBENUMBER
   Else
      Tribes_Processing.Edit
   End If
   Tribes_Processing![SLAVES_OVERSEEN] = "Y"
   Tribes_Processing![Warriors_Assigned] = TActives
   If SLAVERY_LEVEL = 0 Then
      If TActives < (NUMBER_OF_SLAVES / 10) Then
         Tribes_Processing![All_Slaves_Overseen] = "N"
      Else
         Tribes_Processing![All_Slaves_Overseen] = "Y"
      End If
   ElseIf ((((TMouths - TActives) / 10) * SLAVERY_LEVEL) + (TActives * 10)) < NUMBER_OF_SLAVES Then
      Tribes_Processing![All_Slaves_Overseen] = "N"
   Else
      Tribes_Processing![All_Slaves_Overseen] = "Y"
   End If
   Tribes_Processing![Number_Of_Slaves_Overseen] = (((TMouths - TActives) / 10) * SLAVERY_LEVEL) _
   + (TActives * 10)
   Tribes_Processing.UPDATE
   
   Call Check_Turn_Output(",", " oversee slaves", "", 0, "YES")

 
Case "Smoking"
   TRIBE_STATUS = "Perform Smoking"
   ACTIVES_INUSE = TActives
   ' Need to check if have a smokehouse big enough.
   ' Need to check if have enough logs.
     
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "LOG"
       
   TribesGoodsUsage.MoveFirst
   TribesGoodsUsage.Seek "=", TCLANNUMBER, GOODS_TRIBE, "LOG"
       
   If Not TRIBESGOODS.NoMatch Then
      TRIBESGOODS.Edit
      TribesGoodsUsage.Edit
          
      'check maximum actives allocated and no more
      If MAXIMUM_ACTIVES_1 < TActives Then
         TActives = MAXIMUM_ACTIVES_1
      End If
          
      ' check amount of fish against actives allocated
      If TFishing < (TActives * 200) Then
         TActives = CLng(TFishing / 200)
      End If
          
      ' Check available logs against what is required
      If Not TribesGoodsUsage![total_available] >= (TFishing / 10) Then
         TFishing = TribesGoodsUsage![total_available] * 10
      End If
          
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - CLng(TFishing / 10)
      TribesGoodsUsage![Number_Used] = TribesGoodsUsage![Number_Used] + CLng(TFishing / 10)
      TFishing = 0
      TRIBESGOODS.UPDATE
      TribesGoodsUsage.UPDATE
   End If
    
 Case "Specialists"
    TRIBE_STATUS = "Perform Specialists"
  ' MOVE ACTIVES TO TRAINING
   If TItem = "TRAINING" Then
      TribesSpecialists.MoveFirst
      TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, "TRAINING"
       
      If TribesSpecialists.NoMatch Then
         TribesSpecialists.AddNew
         TribesSpecialists![CLAN] = TCLANNUMBER
         TribesSpecialists![TRIBE] = TTRIBENUMBER
         TribesSpecialists![ITEM] = "TRAINING"
         TribesSpecialists![SPECIALISTS] = TActives
         TribesSpecialists![SPECIALISTS_USED] = TActives
         TribesSpecialists![NUMBER_OF_TURNS_TRAINING] = 0
         TribesSpecialists.UPDATE
      Else
         TribesSpecialists.Edit
         TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] + TActives
         TribesSpecialists.UPDATE
      End If
   ElseIf TItem = "Actives" Then
      TribesSpecialists.MoveFirst
      TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, "TRAINING"
  
      If TribesSpecialists.NoMatch Then
         ' need do nothing
      ElseIf TribesSpecialists![NUMBER_OF_TURNS_TRAINING] >= 3 Then
          If TribesSpecialists![SPECIALISTS] <= TActives Then
             TActives = TribesSpecialists![SPECIALISTS]
          End If
          TribesSpecialists.Edit
          TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] - TActives
          TribesSpecialists.UPDATE
          TRIBESINFO.MoveFirst
          TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
          TRIBESINFO.Edit
          TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] + TActives
          TRIBESINFO.UPDATE
       End If
   ElseIf TItem = "Promotion" Then
      TribesSpecialists.MoveFirst
      TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, "TRAINING"
     
      If TribesSpecialists.NoMatch Then
         ' need do nothing
      Else
         If TribesSpecialists![SPECIALISTS] < TActives Then
            TActives = TribesSpecialists![SPECIALISTS]
         End If
         TribesSpecialists.Edit
         TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] - TActives
         TribesSpecialists.UPDATE
         TribesSpecialists.MoveFirst
         TribesSpecialists.Seek "=", TCLANNUMBER, TTRIBENUMBER, TDistinction
         If TribesSpecialists.NoMatch Then
            TribesSpecialists.AddNew
            TribesSpecialists![CLAN] = TCLANNUMBER
            TribesSpecialists![TRIBE] = TTRIBENUMBER
            TribesSpecialists![ITEM] = TDistinction
            TribesSpecialists![SPECIALISTS] = TActives
            TribesSpecialists![SPECIALISTS_USED] = TActives
            TribesSpecialists.UPDATE
         Else
            TribesSpecialists.Edit
            TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] + TActives
            TribesSpecialists.UPDATE
         End If
      End If
   End If
    
Case "STONEWORK"
   TRIBE_STATUS = "Perform Stonework"
   TACLAN = TCLANNUMBER
   TATRIBE = TTRIBENUMBER
   TAACTIVITY = TActivity
   TAITEM = TItem
   TAACTIVES = TActives
  
   If TSkillok(1) = "Y" Then
      If Left(TItem, 4) = "KILN" Then
         Call Calc_Engineering(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TAACTIVES)
      ElseIf Left(TItem, 4) = "OVEN" Then
         Call Calc_Engineering(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TAACTIVES)
      ElseIf Left(TItem, 7) = "SMELTER" Then
         Call Calc_Engineering(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TAACTIVES)
      ElseIf Left(TItem, 6) = "BURNER" Then
         Call Calc_Engineering(TACLAN, TATRIBE, GOODS_TRIBE, TAACTIVITY, TAITEM, TADISTINCTION, TAACTIVES)
      Else
         Call PERFORM_STONEWORK
      End If
   Else
      Call Check_Turn_Output(",", " No Stonework Skill", "", 0, "NO")
   End If

Case "SUPPRESSION"
   TRIBE_STATUS = "Perform Suppression"
   ' NEED TO IDENTIFY WHAT ITEMS ARE BEING USED
     
Case "TANNING"
   TRIBE_STATUS = "Perform Tanning"
   If TSkillok(1) = "Y" Then
      Call PERFORM_TANNING
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Tanning Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Tanning Skill", "", 0, "NO")
      End If
   End If

Case "TRIBE DEFENCE"
   ' NO ACTION REQUIRED
    
Case "WAXWORK"
   TRIBE_STATUS = "Perform Waxwork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_WAXWORK
   Else
      If Right(TurnActOutPut, 3) = "^B " Then
         Call Check_Turn_Output("", " No Waxwork Skill", "", 0, "NO")
      Else
         Call Check_Turn_Output(",", " No Waxwork Skill", "", 0, "NO")
      End If
   End If

Case "WEAPONS"
   TRIBE_STATUS = "Perform Weapons"
   If TSkillok(1) = "Y" Then
      If TItem = "Staves" Then
         Select Case TRIBES_TERRAIN
         Case "DECIDUOUS"
            Call PERFORM_WEAPONS
         Case "DECIDUOUS FLAT"
            Call PERFORM_WEAPONS
         Case "DECIDUOUS FOREST"
            Call PERFORM_WEAPONS
         Case "DECIDUOUS HILLS"
            Call PERFORM_WEAPONS
         Case "JUNGLE"
            Call PERFORM_WEAPONS
         Case "JUNGLE HILLS"
            Call PERFORM_WEAPONS
         Case "LOW JUNGLE MOUNTAINS"
            Call PERFORM_WEAPONS
         Case "LOW JUNGLE MT"
            Call PERFORM_WEAPONS
         Case Else
            Call Check_Turn_Output(",", " Invalid terrain for staves, ", "", 0, "NO")
         End Select
      Else
         Call PERFORM_WEAPONS
      End If
   Else
      Call Check_Turn_Output(",", " No Weapons Skill", "", 0, "NO")
   End If

Case "WEAVING"
   TRIBE_STATUS = "Perform Weaving"
   If TSkillok(1) = "Y" Then
      Call PERFORM_WEAVING
   Else
      Call Check_Turn_Output(",", " No Weaving Skill", "", 0, "NO")
   End If

Case "Whaling"
   TRIBE_STATUS = "Perform Whaling"
   ' prompt for whales caught
   DICE1 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)
   If DICE1 < (20 + WHALING_LEVEL) Then
      Whale = InputBox("What size of WHALE has been caught? (S/M/L)", "Size", "N")
      NumWhales = InputBox("How many whales caught?", "NUMBER", "1")
      If Whale = "S" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "whale - small"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.AddNew
            TRIBESGOODS![CLAN] = TCLANNUMBER
            TRIBESGOODS![TRIBE] = GOODS_TRIBE
            TRIBESGOODS![ITEM_TYPE] = "RAW"
            TRIBESGOODS![ITEM] = "whale - small"
            TRIBESGOODS![ITEM_NUMBER] = NumWhales
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + NumWhales
         End If
      ElseIf Whale = "M" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "whale - medium"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.AddNew
            TRIBESGOODS![CLAN] = TCLANNUMBER
            TRIBESGOODS![TRIBE] = GOODS_TRIBE
            TRIBESGOODS![ITEM_TYPE] = "RAW"
            TRIBESGOODS![ITEM] = "whale - medium"
            TRIBESGOODS![ITEM_NUMBER] = NumWhales
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + NumWhales
         End If
      Else
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "RAW", "whale - large"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.AddNew
            TRIBESGOODS![CLAN] = TCLANNUMBER
            TRIBESGOODS![TRIBE] = GOODS_TRIBE
            TRIBESGOODS![ITEM_TYPE] = "RAW"
            TRIBESGOODS![ITEM] = "whale - large"
            TRIBESGOODS![ITEM_NUMBER] = NumWhales
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![NUMBER] = TRIBESGOODS![NUMBER] + NumWhales
         End If
      End If
      TRIBESGOODS.UPDATE
      DoCmd.Hourglass True
   End If
   ' Update output line
   TurnActOutPut = TurnActOutPut & ", Caught " & NumWhales
   If Whale = "S" Then
      TurnActOutPut = TurnActOutPut & " S/Whales "
   ElseIf Whale = "M" Then
      TurnActOutPut = TurnActOutPut & " M/Whales "
   ElseIf Whale = "L" Then
      TurnActOutPut = TurnActOutPut & " L/Whales "
   Else
      TurnActOutPut = TurnActOutPut & " 0 Whales "
   End If
      
Case "WOODWORK"
   TRIBE_STATUS = "Perform Woodwork"
   If TSkillok(1) = "Y" Then
      Call PERFORM_WOODWORK
   Else
      Call Check_Turn_Output(",", " No Woodwork Skill", "", 0, "NO")
   End If

Case Else
   '?????????????? - means activity is not catered for
      Call Check_Turn_Output(",", " TActivity not catered for", "", 0, "NO")
End Select

ERR_A500_MAIN_PROCESS_CLOSE:
   Exit Function

ERR_A500_MAIN_PROCESS:
   Call A999_ERROR_HANDLING
   Resume ERR_A500_MAIN_PROCESS_CLOSE

End Function

Public Function A999_ERROR_HANDLING()
  Msg = "The Process " & TRIBE_STATUS & "has received the following error message "
  Msg = Msg & Chr(13) & Chr(10) & "Error # " & Err & " " & Error$ & Chr(13) & Chr(10)
  Msg = Msg & "The Clan was " & TCLANNUMBER & " The Tribe was " & TTRIBENUMBER
  Msg = Msg & Chr(13) & Chr(10) & " GIVE THIS INFO TO Jeff."
  MsgBox (Msg)


End Function

Public Function A800_OUTPUT_PROCESSING()
On Error GoTo ERR_A800_OUTPUT_PROCESSING
TRIBE_STATUS = "A800_OUTPUT_PROCESSING"
DebugOP "A800_OUTPUT_PROCESSING"

If Len(TurnActOutPut) > 0 Then
   Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TTRIBENUMBER, "ACTIVITIES", 1, TurnActOutPut, "No")
End If

ERR_A800_OUTPUT_PROCESSING_CLOSE:
   Exit Function

ERR_A800_OUTPUT_PROCESSING:
   Call A999_ERROR_HANDLING
   Resume ERR_A800_OUTPUT_PROCESSING_CLOSE

End Function

Public Function A150_Open_Tables(WHICH_TABLE)
On Error GoTo ERR_OPEN_TABLES

TRIBE_STATUS = "A150_Open_Tables"
DebugOP "A150_Open_Tables"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "ACTIVITIES" Then
   Set ActivitiesTable = TVDB.OpenRecordset("Activities")
   ActivitiesTable.index = "PRIMARYKEY"
   ActivitiesTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Activity" Then
   Set ItemsTable = TVDB.OpenRecordset("Activity")
   ItemsTable.index = "SECONDARYKEY"
   ItemsTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "ACTIVITY_SEQUENCE" Then
   Set ACTSEQTAB = TVDB.OpenRecordset("ACTIVITY_SEQUENCE")
   ACTSEQTAB.index = "SECONDARYKEY"
   ACTSEQTAB.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "COMPLETED_RESEARCH" Then
   Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
   COMPRESTAB.index = "PRIMARYKEY"
   COMPRESTAB.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "global" Then
   Set globalinfo = TVDBGM.OpenRecordset("GLOBAL")
   globalinfo.index = "PRIMARYKEY"
   globalinfo.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "GAMES_WEATHER" Then
   Set GAMES_WEATHER = TVDBGM.OpenRecordset("GAMES_WEATHER")
   GAMES_WEATHER.index = "PRIMARYKEY"
   GAMES_WEATHER.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEX_MAP" Then
   Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
   hexmaptable.index = "PRIMARYKEY"
   hexmaptable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEX_MAP_CONST" Then
   Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
   HEXMAPCONST.index = "FORTHKEY"
   HEXMAPCONST.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Building_Used" Then
   Set Building_Used = TVDB.OpenRecordset("Building_Usage")
   Building_Used.index = "PRIMARYKEY"
   Building_Used.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEX_MAP_MINERALS" Then
   Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
   HEXMAPMINERALS.index = "PRIMARYKEY"
   HEXMAPMINERALS.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEX_MAP_POLITICS" Then
   Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_POLITICS")
   HEXMAPPOLITICS.index = "PRIMARYKEY"
   HEXMAPPOLITICS.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEXMAP_FARMING" Then
   Set FarmingTable = TVDBGM.OpenRecordset("HEXMAP_FARMING")
   FarmingTable.index = "PRIMARYKEY"
   FarmingTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "HEXMAP_PERMANENT_FARMING" Then
   Set PermFarmingTable = TVDBGM.OpenRecordset("HEXMAP_PERMANENT_FARMING")
   PermFarmingTable.index = "PRIMARYKEY"
   PermFarmingTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "IMPLEMENT_USAGE" Then
   Set ImplementUsage = TVDB.OpenRecordset("IMPLEMENT_USAGE")
   ImplementUsage.index = "PRIMARYKEY"
   ImplementUsage.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "IMPLEMENTS" Then
   Set ImplementsTable = TVDB.OpenRecordset("IMPLEMENTS")
   ImplementsTable.index = "PRIMARYKEY"
   ImplementsTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "MODIFIERS" Then
   Set MODTABLE = TVDBGM.OpenRecordset("MODIFIERS")
   MODTABLE.index = "PRIMARYKEY"
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "PACIFICATION_TABLE" Then
   Set PACIFICATION_TABLE = TVDBGM.OpenRecordset("PACIFICATION_TABLE")
   PACIFICATION_TABLE.index = "PRIMARYKEY"
   PACIFICATION_TABLE.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Process_Tribes_Activity" Then
   Set PROCESSACTIVITY = TVDBGM.OpenRecordset("Process_Tribes_Activity")
   PROCESSACTIVITY.index = "PRIMARYKEY"
   PROCESSACTIVITY.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Process_Tribes_Item_Allocation" Then
   Set PROCESSITEMS = TVDBGM.OpenRecordset("Process_Tribes_Item_Allocation")
   PROCESSITEMS.index = "PRIMARYKEY"
   PROCESSITEMS.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "SEASON" Then
   Set SEASONTABLE = TVDB.OpenRecordset("SEASON")
   SEASONTABLE.index = "PRIMARYKEY"
   SEASONTABLE.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "SEEKING_RETURNS" Then
   Set SeekingReturnsTable = TVDBGM.OpenRecordset("SEEKING_RETURNS_TABLE")
   SeekingReturnsTable.index = "PRIMARYKEY"
   SeekingReturnsTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "SKILLS" Then
   Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
   SKILLSTABLE.index = "PRIMARYKEY"
   SKILLSTABLE.MoveFirst
End If

If WHICH_TABLE = "ALL" Then
   Set TrAct_Req_Later = TVDB.OpenRecordset("Tribe_Activity_Required_By_Later_Activities")
   TrAct_Req_Later.index = "PRIMARYKEY"
   TrAct_Req_Later.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Tribes_General_Info" Then
   Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
   TRIBESINFO.index = "PRIMARYKEY"
   TRIBESINFO.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Tribes_Goods" Then
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Tribes_Goods_Usage" Then
   Set TribesGoodsUsage = TVDB.OpenRecordset("Tribes_Goods_Usage")
   TribesGoodsUsage.index = "PRIMARYKEY"
   TribesGoodsUsage.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Tribes_Processing" Then
   Set Tribes_Processing = TVDBGM.OpenRecordset("Tribes_Processing")
   Tribes_Processing.index = "PRIMARYKEY"
   Tribes_Processing.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Tribes_Specialists" Then
   Set TribesSpecialists = TVDBGM.OpenRecordset("Tribes_Specialists")
   TribesSpecialists.index = "PRIMARYKEY"
   TribesSpecialists.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "TURNS_ACTIVITIES" Then
   Set OutPutTable = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
   OutPutTable.index = "PRIMARYKEY"
End If

If WHICH_TABLE = "ALL" Then
   Set Turn_Info_Req_NxTurn = TVDBGM.OpenRecordset("Turn_Info_Reqd_Next_Turn")
   Turn_Info_Req_NxTurn.index = "PRIMARYKEY"
   Turn_Info_Req_NxTurn.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "Under_Construction" Then
   Set ConstructionTable = TVDBGM.OpenRecordset("Under_Construction")
   ConstructionTable.index = "PRIMARYKEY"
   ConstructionTable.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_ANIMALS" Then
   Set VALIDANIMALS = TVDB.OpenRecordset("VALID_ANIMALS")
   VALIDANIMALS.index = "primarykey"
   VALIDANIMALS.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_BUILDINGS" Then
   Set VALID_CONST = TVDB.OpenRecordset("VALID_BUILDINGS")
   VALID_CONST.index = "primarykey"
   VALID_CONST.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_GOODS" Then
   Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
   VALIDGOODS.index = "primarykey"
   VALIDGOODS.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_MINERALS" Then
   Set VALIDMINERALS = TVDB.OpenRecordset("VALID_MINERALS")
   VALIDMINERALS.index = "primarykey"
   VALIDMINERALS.MoveFirst
End If
  
If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_SHIPS" Then
   Set VALIDSHIPS = TVDB.OpenRecordset("VALID_SHIPS")
   VALIDSHIPS.index = "PRIMARYKEY"
   VALIDSHIPS.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_TERRAIN" Then
   Set TERRAINTABLE = TVDB.OpenRecordset("VALID_TERRAIN")
   TERRAINTABLE.index = "PRIMARYKEY"
   TERRAINTABLE.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "VALID_WIND" Then
   Set WINDTABLE = TVDB.OpenRecordset("VALID_WIND")
   WINDTABLE.index = "PRIMARYKEY"
   WINDTABLE.MoveFirst
End If

If WHICH_TABLE = "ALL" Or WHICH_TABLE = "WEATHER" Then
   Set WEATHERTABLE = TVDBGM.OpenRecordset("WEATHER")
   WEATHERTABLE.index = "PRIMARYKEY"
End If
    DebugOP "A150_Open_Tables - completed"
ERR_OPEN_TABLES_CLOSE:
   Exit Function

ERR_OPEN_TABLES:
If (Err = 3021) Then  ' 3021 = No Current Record
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_OPEN_TABLES_CLOSE
End If


End Function

Public Function A900_Close_Tables()
On Error GoTo ERR_A900_Close_Tables
TRIBE_STATUS = "A900_Close_Tables"
DebugOP "A900_Close_Tables"
   ActivitiesTable.Close
   ACTSEQTAB.Close
   COMPRESTAB.Close
   ConstructionTable.Close
   FarmingTable.Close
   GAMES_WEATHER.Close
   globalinfo.Close
   hexmaptable.Close
   HEXMAPCONST.Close
   Building_Used.Close
   HEXMAPMINERALS.Close
   HEXMAPPOLITICS.Close
   ItemsTable.Close
   ImplementUsage.Close
   ImplementsTable.Close
   MODTABLE.Close
   OutPutTable.Close
   PermFarmingTable.Close
   PROCESSACTIVITY.Close
   PROCESSITEMS.Close
   SEASONTABLE.Close
   SeekingReturnsTable.Close
   SKILLSTABLE.Close
   TERRAINTABLE.Close
   TrAct_Req_Later.Close
   TRIBESINFO.Close
   TRIBESGOODS.Close
   TribesGoodsUsage.Close
   Tribes_Processing.Close
   TribesSpecialists.Close
   Turn_Info_Req_NxTurn.Close
   VALIDGOODS.Close
   WINDTABLE.Close
   WEATHERTABLE.Close
   GMTABLE.Close
   TEMPCOMPRESTAB.Close
   CROP_TABLE.Close
   CLIMATETABLE.Close
   Goods_Tribes_Processed.Close
   PopTable.Close
   PROVS_AVAIL_TABLE.Close
   RESEARCHTABLE.Close
   TRIBESBOOKS.Close
   VALIDANIMALS.Close
   VALIDBUILDINGS.Close
   VALID_CONST.Close
   VALIDMINERALS.Close
   VALIDSHIPS.Close
   
ERR_A900_Close_Tables_CLOSE:
   Exit Function

ERR_A900_Close_Tables:
   Resume Next

End Function

Public Function A250_Get_Tribe_Info()
On Error GoTo ERR_A250_Get_Tribe_Info
TRIBE_STATUS = "Get Tribe Info Data"
DebugOP "A250_Get_Tribe_Info"

INVALID_TRIBE = "N"
' Get the Number of Warriors, Actives & Slaves available
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
  
If TRIBESINFO.NoMatch Then
   Msg = "Clan : " & TCLANNUMBER & " Party : " & TTRIBENUMBER & " " & LINEFEED
   Msg = Msg & " is not on tribes table.  Needs to have 'NEW TRIBE' form updated."
   MsgBox (Msg)
   TurnActOutPut = TurnActOutPut & TTRIBENUMBER & " was invalid and the activity was not processed, "
   INVALID_TRIBE = "Y"
Else
  TActivesAvailable = TRIBESINFO![WARRIORS] + TRIBESINFO![ACTIVES]
  If Not IsNull(TRIBESINFO![SLAVE]) Then
     TActivesAvailable = TActivesAvailable + TRIBESINFO![SLAVE]
  End If
  If Not IsNull(TRIBESINFO![HIRELINGS]) Then
     TActivesAvailable = TActivesAvailable + TRIBESINFO![HIRELINGS]
  End If
  If Not IsNull(TRIBESINFO![LOCALS]) Then
     TActivesAvailable = TActivesAvailable + TRIBESINFO![LOCALS]
  End If
  If Not IsNull(TRIBESINFO![Auxiliaries]) Then
     TActivesAvailable = TActivesAvailable + TRIBESINFO![Auxiliaries]
  End If
   
  TInActives = TRIBESINFO![INACTIVES]
  TMouths = TActivesAvailable
  TMouths = TMouths + TRIBESINFO![INACTIVES]
  If Not IsNull(TRIBESINFO![MERCENARIES]) Then
     TMouths = TMouths + TRIBESINFO![MERCENARIES]
  End If
  
  RESEARCH_FOUND = "N"

  Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Inactive Workers")

  If RESEARCH_FOUND = "Y" Then
     TActivesAvailable = TActivesAvailable + CLng(TInActives / 3)
  End If
  
  NUMBER_OF_SLAVES = TRIBESINFO![SLAVE]
  
  TRIBESGOODS.index = "PRIMARYKEY"
  TRIBESGOODS.MoveFirst
  TRIBESGOODS.Seek "=", TCLANNUMBER, TTRIBENUMBER, "ANIMAL", "HERDING DOG"
  If Not TRIBESGOODS.NoMatch Then
     TMouths = TMouths + CLng(TRIBESGOODS![ITEM_NUMBER] / 2)
  End If
  
  If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
     GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
  Else
     GOODS_TRIBE = TTRIBENUMBER
  End If

  Tribes_Current_Hex = TRIBESINFO![CURRENT HEX]
  Meeting_House_Hex = TRIBESINFO![CURRENT HEX]
  
  TRIBESINFO.MoveFirst
  TRIBESINFO.Seek "=", TCLANNUMBER, GOODS_TRIBE
  Goods_Tribes_Current_Hex = TRIBESINFO![CURRENT HEX]
  
  TRIBESINFO.MoveFirst
  TRIBESINFO.Seek "=", TOwning_Clan, TOwning_Tribe
  CONST_Tribes_Current_Hex = TRIBESINFO![CURRENT HEX]
  
 ' If Not Tribes_Current_Hex = CONST_Tribes_Current_Hex Then
 '       Msg = "Clan : " & TCLANNUMBER & " Party : " & TTRIBENUMBER & " " & LINEFEED
 '       Msg = Msg & " OWNER_TRIBE " & TOwning_Tribe & " is not located in the same hex."
 '       MsgBox (Msg)
 '       TurnActOutPut = TurnActOutPut & TTRIBENUMBER & " OWNING_TRIBE was invalid (not located at the same place)  and activity was not processed, "
 '       INVALID_TRIBE = "Y"
 ' End If
  
  TRIBESINFO.MoveFirst
  TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
  
  ' Find out the hex of the tribe with the meeting house.  Either the tribe or the goods tribe
  HEXMAPCONST.index = "FORTHKEY"
  HEXMAPCONST.Seek "=", CONST_Tribes_Current_Hex, TOwning_Clan, "MEETING HOUSE"

  If HEXMAPCONST.NoMatch Then
     TRIBESINFO.MoveFirst
     TRIBESINFO.Seek "=", TCLANNUMBER, GOODS_TRIBE
     CONST_MEETING_HOUSE_FOUND = "N"
  Else
     CONST_MEETING_HOUSE_FOUND = "Y"
  End If
 
' Find out the hex of the tribe with the meeting house.  Either the tribe or the goods tribe
  HEXMAPCONST.index = "FORTHKEY"
  HEXMAPCONST.Seek "=", Tribes_Current_Hex, TCLANNUMBER, "MEETING HOUSE"

  If HEXMAPCONST.NoMatch Then
     TRIBESINFO.MoveFirst
     TRIBESINFO.Seek "=", TCLANNUMBER, GOODS_TRIBE
     MEETING_HOUSE_FOUND = "N"
  Else
     MEETING_HOUSE_FOUND = "Y"
  End If
 
  If HEXMAPCONST.NoMatch Then
     ' Reset back to what the code is expecting
     TRIBESINFO.MoveFirst
     TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
     MEETING_HOUSE_FOUND = "N"
  Else
     MEETING_HOUSE_FOUND = "Y"
  End If
 
  TRIBES_TERRAIN = TRIBESINFO![CURRENT TERRAIN]
  
  Government_Level = TRIBESINFO![GOVT LEVEL]
  
  If Not IsNull(TRIBESINFO![RELIGION]) Then
     TRIBES_RELIGION = TRIBESINFO![RELIGION]
  Else
     TRIBES_RELIGION = "NONE"
  End If
  If Not IsNull(TRIBESINFO![CULT]) Then
     TRIBES_CULT = TRIBESINFO![CULT]
  Else
     TRIBES_CULT = "NONE"
  End If
  If Not IsNull(TRIBESINFO![POP TRIBE]) Then
     TRIBES_POP_TRIBE = TRIBESINFO![POP TRIBE]
  Else
     TRIBES_POP_TRIBE = TTRIBENUMBER
  End If
  If Not IsNull(TRIBESINFO![COST CLAN]) Then
     TRIBES_COST_CLAN = TRIBESINFO![COST CLAN]
  Else
     TRIBES_COST_CLAN = TCLANNUMBER
  End If
  If Not IsNull(TRIBESINFO![MORALE]) Then
     TRIBES_MORALE = TRIBESINFO![MORALE]
  Else
     TRIBES_MORALE = 1
  End If
  
  Call GET_MODIFIERS
  

End If
ERR_A250_Get_Tribe_Info_CLOSE:
   Exit Function


ERR_A250_Get_Tribe_Info:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Tribe_Info_CLOSE
End If
End Function

Public Function A250_Get_Hex_Info()
On Error GoTo ERR_A250_Get_Hex_Info
TRIBE_STATUS = "GET_HEX_MAP_DATA"
DebugOP "A250_Get_Hex_Info"

hexmaptable.MoveFirst
hexmaptable.Seek "=", Tribes_Current_Hex

If hexmaptable.NoMatch Then
   Msg = "HEX OF TRIBE NOT FOUND"
   MsgBox (Msg)
Else
   FARMING_TERRAIN = hexmaptable![TERRAIN]
   FRESH_WATER = hexmaptable![FRESH WATER]
   QUARRYING = hexmaptable![QUARRYING]
   ROAMING_HERD = hexmaptable![ROAMING HERD]
   SALMON_RUN = hexmaptable![SALMON RUN]
   FISH_AREA = hexmaptable![FISH AREA]
   WHALE_AREA = hexmaptable![WHALE AREA]
   RIVER_N = Mid(hexmaptable![Borders], 1, 2)
   RIVER_NE = Mid(hexmaptable![Borders], 3, 2)
   RIVER_SE = Mid(hexmaptable![Borders], 5, 2)
   RIVER_S = Mid(hexmaptable![Borders], 7, 2)
   RIVER_SW = Mid(hexmaptable![Borders], 9, 2)
   RIVER_NW = Mid(hexmaptable![Borders], 11, 2)
   If Not IsNull(hexmaptable![WEATHER_ZONE]) Then
      Hexmaps_WEATHER_ZONE = hexmaptable![WEATHER_ZONE]
      CURRENT_WEATHER_ZONE = hexmaptable![WEATHER_ZONE]
   Else
      Hexmaps_WEATHER_ZONE = "NULL"
      CURRENT_WEATHER_ZONE = "NULL"
   End If
End If
  
HEXMAPPOLITICS.MoveFirst
HEXMAPPOLITICS.Seek "=", Tribes_Current_Hex

If HEXMAPPOLITICS.NoMatch Then
   'MSG = "HEX OF TRIBE NOT FOUND"
   'MsgBox (MSG)
   CURRENT_HEX_POP = 0
   CURRENT_HEX_PAC_LEV = 0
Else
   If IsNull(HEXMAPPOLITICS![POPULATION]) Then
      CURRENT_HEX_POP = 0
   Else
      CURRENT_HEX_POP = HEXMAPPOLITICS![POPULATION]
   End If
   If IsNull(HEXMAPPOLITICS![PACIFICATION_LEVEL]) Then
      CURRENT_HEX_PAC_LEV = 0
   Else
      CURRENT_HEX_PAC_LEV = HEXMAPPOLITICS![PACIFICATION_LEVEL]
   End If
End If

ERR_A250_Get_Hex_Info_CLOSE:
   Exit Function


ERR_A250_Get_Hex_Info:
If (Err = 3021) Then          ' NO CURRENT RECORD
   Resume Next
   
Else
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Hex_Info_CLOSE
End If
End Function

Public Function A250_Get_Weather_Info()
On Error GoTo ERR_A250_Get_Weather_Info
TRIBE_STATUS = "A250_Get_Weather_Info"
DebugOP "A250_Get_Weather_Info"

  If CURRENT_WEATHER_ZONE = "GREEN" Then
     CURRENT_WEATHER = globalinfo![Zone1]
     CURRENT_WIND = globalinfo![Wind1]
     CURRENT_CLIMATE = globalinfo![CLIMATE 1]
  ElseIf CURRENT_WEATHER_ZONE = "RED" Then
     CURRENT_WEATHER = globalinfo![Zone2]
     CURRENT_WIND = globalinfo![Wind2]
     CURRENT_CLIMATE = globalinfo![CLIMATE 2]
  ElseIf CURRENT_WEATHER_ZONE = "ORANGE" Then
     CURRENT_WEATHER = globalinfo![Zone3]
     CURRENT_WIND = globalinfo![Wind3]
     CURRENT_CLIMATE = globalinfo![CLIMATE 3]
  ElseIf CURRENT_WEATHER_ZONE = "YELLOW" Then
     CURRENT_WEATHER = globalinfo![Zone4]
     CURRENT_WIND = globalinfo![Wind4]
     CURRENT_CLIMATE = globalinfo![CLIMATE 4]
  ElseIf CURRENT_WEATHER_ZONE = "BLUE" Then
     CURRENT_WEATHER = globalinfo![Zone5]
     CURRENT_WIND = globalinfo![Wind5]
     CURRENT_CLIMATE = globalinfo![CLIMATE 5]
  ElseIf CURRENT_WEATHER_ZONE = "BROWN" Then
     CURRENT_WEATHER = globalinfo![Zone6]
     CURRENT_WIND = globalinfo![Wind6]
     CURRENT_CLIMATE = globalinfo![CLIMATE 6]
  End If

If FISH_AREA = "Y" Then
   SEASON_FISHING = 0.2
End If
If SALMON_RUN = "Y" Then
   SEASON_FISHING = 0.2
End If

WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "HERDING", "1"
WEATHER_HERDING_GROUP_1 = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "HERDING", "2"
WEATHER_HERDING_GROUP_2 = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "HERDING", "3"
WEATHER_HERDING_GROUP_3 = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "HUNTING", "PROVS"
HUNTING_WEATHER = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "MINING", "ORE"
MINING_WEATHER = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "MINING", "ACCIDENTS"
MINING_ACCIDENTS_WEATHER = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "APIARISM", "HONEY"
HONEY_WEATHER = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "APIARISM", "WAX"
WAX_WEATHER = WEATHERTABLE![Modifier]
WEATHERTABLE.MoveFirst
WEATHERTABLE.Seek "=", CURRENT_WEATHER, "FISHING", "PROVS"
FISHING_WEATHER = WEATHERTABLE![Modifier]

If CURRENT_SEASON = "SPRING" Then
   If RIVER_N = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   ElseIf RIVER_NE = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   ElseIf RIVER_SE = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   ElseIf RIVER_S = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   ElseIf RIVER_SW = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   ElseIf RIVER_NW = "RI" Then
      SEASON_FISHING = SEASON_FISHING + 1.8
   End If
End If

ERR_A250_Get_Weather_Info_CLOSE:
   Exit Function


ERR_A250_Get_Weather_Info:
   Call A999_ERROR_HANDLING
   Resume ERR_A250_Get_Weather_Info_CLOSE

End Function


Function CALC_BUILDING_ADDITION()
On Error GoTo ERR_CALC_BUILDING_ADDITION
TRIBE_STATUS = "CALC_BUILDING_ADDITION"

' *************************************************************
' MAXIMUM OF TEN (10) BUILDINGS eg Only allowed 10 Refineries.
' *************************************************************

Dim TINCREASE As Long
Dim Only_Update_Field_1 As String
Dim sResult As String
TPOSITION = 0



BRACKET = InStr(CONSTRUCTION, "(")
If BRACKET > 0 Then
   NEWCONSTRUCTION = Left(CONSTRUCTION, (BRACKET - 1))
Else
   NEWCONSTRUCTION = CONSTRUCTION
End If

If isContainerBuilding(CONSTRUCTION) Then
    sResult = sAddContainerBuilding(CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, NEWCONSTRUCTION)
    Exit Function
End If

If Not sContainerBuilding(CONSTRUCTION) = "FALSE" Then
    sResult = sAddInstallationConstruction(CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, NEWCONSTRUCTION, TBuilding)
    Exit Function
End If

'If NEWCONSTRUCTION = "WOODEN TOWER" Then
'   NEWCONSTRUCTION = "TOWER WOODEN"
'End If

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, NEWCONSTRUCTION

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = CONST_Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = NEWCONSTRUCTION
   HEXMAPCONST.UPDATE
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, NEWCONSTRUCTION
End If

HEXMAPCONST.Edit

VALID_CONST.Seek "=", NEWCONSTRUCTION

If Not VALID_CONST.NoMatch Then
   If Left(CONSTRUCTION, 6) = "BURNER" Then
      HEXMAPCONST.index = "PRIMARYKEY"
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, NEWCONSTRUCTION
      HEXMAPCONST.Delete
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", CONST_Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "CHARHOUSE"
      If HEXMAPCONST.NoMatch Then
        MsgBox "No match found in HEXMAPCONST table for Tribe: " & CONSTTRIBE & ", Item: " & "CHARHOUSE", vbExclamation, "No Match Error"
      End If
      HEXMAPCONST.Edit
      TPOSITION = TBuilding
   
   ElseIf Left(CONSTRUCTION, 4) = "KEEP" Then
      If IsNull(HEXMAPCONST![KEEP]) Then
         HEXMAPCONST![KEEP] = "1"
      Else
         HEXMAPCONST![KEEP] = HEXMAPCONST![KEEP] & ",1"
      End If
   
   ElseIf Left(CONSTRUCTION, 12) = "WOODEN TOWER" Then
      TPOSITION = 1
   
   ElseIf Left(CONSTRUCTION, 12) = "STONE TOWER" Then
      TPOSITION = 1
   
   ElseIf Left(CONSTRUCTION, 4) = "WELL" Then
      TPOSITION = 1
   
   ElseIf Left(CONSTRUCTION, 9) = "RUNE ROAD" Then
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", CONST_Tribes_Current_Hex

      HEX_N = GET_MAP_NORTH(CONST_Tribes_Current_Hex)
      HEX_NE = GET_MAP_NORTH_EAST(CONST_Tribes_Current_Hex)
      HEX_SE = GET_MAP_SOUTH_EAST(CONST_Tribes_Current_Hex)
      HEX_S = GET_MAP_SOUTH(CONST_Tribes_Current_Hex)
      HEX_SW = GET_MAP_SOUTH_WEST(CONST_Tribes_Current_Hex)
      HEX_NW = GET_MAP_NORTH_WEST(CONST_Tribes_Current_Hex)
   
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_N
   
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_N = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
         ROAD_N = "Y"
      Else
         ROAD_N = "R"
      End If
   
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_NE
   
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_NE = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
        ROAD_NE = "Y"
      Else
         ROAD_NE = "R"
      End If

      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_SE
      
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_SE = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
         ROAD_SE = "Y"
      Else
         ROAD_SE = "R"
      End If
   
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_S
      
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_S = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
         ROAD_S = "Y"
      Else
         ROAD_S = "R"
      End If

      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_SW
      
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_SW = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
         ROAD_SW = "Y"
      Else
          ROAD_SW = "R"
      End If
   
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", HEX_NW
      
      If hexmaptable![TERRAIN] = "OCEAN" Then
         ROAD_NW = "N"
      ElseIf hexmaptable![TERRAIN] = "LAKE" Then
         ROAD_NW = "Y"
      Else
         ROAD_NW = "R"
      End If

      hexmaptable.MoveFirst
      hexmaptable.Seek "=", CONST_Tribes_Current_Hex
      hexmaptable.Edit
      hexmaptable![ROADS] = ROAD_N & ROAD_NE & ROAD_SE & ROAD_S & ROAD_SW & ROAD_NW
      hexmaptable.UPDATE
   End If
 
Else
    'MsgBox (CONSTRUCTION & " not included in program")
End If



If TPOSITION > 0 Then
   If TPOSITION = 1 Then
      HEXMAPCONST![1] = HEXMAPCONST![1] + 1
   ElseIf TPOSITION = 2 Then
      HEXMAPCONST![2] = HEXMAPCONST![2] + 1
   ElseIf TPOSITION = 3 Then
      HEXMAPCONST![3] = HEXMAPCONST![3] + 1
   ElseIf TPOSITION = 4 Then
      HEXMAPCONST![4] = HEXMAPCONST![4] + 1
   ElseIf TPOSITION = 5 Then
      HEXMAPCONST![5] = HEXMAPCONST![5] + 1
   ElseIf TPOSITION = 6 Then
      HEXMAPCONST![6] = HEXMAPCONST![6] + 1
   ElseIf TPOSITION = 7 Then
      HEXMAPCONST![7] = HEXMAPCONST![7] + 1
   ElseIf TPOSITION = 8 Then
      HEXMAPCONST![8] = HEXMAPCONST![8] + 1
   ElseIf TPOSITION = 9 Then
      HEXMAPCONST![9] = HEXMAPCONST![9] + 1
   ElseIf TPOSITION = 10 Then
      HEXMAPCONST![10] = HEXMAPCONST![10] + 1
   End If

Else
If VALID_CONST![Update_Field_1] = "Y" Then
   HEXMAPCONST![1] = HEXMAPCONST![1] + 1
ElseIf HEXMAPCONST![1] = 0 Then
   HEXMAPCONST![1] = 1
ElseIf HEXMAPCONST![2] = 0 Then
   HEXMAPCONST![2] = 1
ElseIf HEXMAPCONST![3] = 0 Then
   HEXMAPCONST![3] = 1
ElseIf HEXMAPCONST![4] = 0 Then
   HEXMAPCONST![4] = 1
ElseIf HEXMAPCONST![5] = 0 Then
   HEXMAPCONST![5] = 1
ElseIf HEXMAPCONST![6] = 0 Then
   HEXMAPCONST![6] = 1
ElseIf HEXMAPCONST![7] = 0 Then
   HEXMAPCONST![7] = 1
ElseIf HEXMAPCONST![8] = 0 Then
   HEXMAPCONST![8] = 1
ElseIf HEXMAPCONST![9] = 0 Then
   HEXMAPCONST![9] = 1
ElseIf HEXMAPCONST![10] = 0 Then
   HEXMAPCONST![10] = 1
Else
   'MsgBox (CONSTRUCTION & " - this tribe has reached its maximum of 10 buildings")
End If
End If

HEXMAPCONST.UPDATE

ERR_CALC_BUILDING_ADDITION_CLOSE:
   Exit Function

ERR_CALC_BUILDING_ADDITION:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_BUILDING_ADDITION_CLOSE

End Function
Public Function CALC_CLOTH_USED(CLOTHUSED)
On Error GoTo ERR_CALC_CLOTH_USED
TRIBE_STATUS = "CALC_CLOTH_USED"

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "CLOTH"

If TOTAL_CLOTH > CLOTHUSED Then
   Continue = "Y"
Else
   Continue = "N"
End If

If TRIBESGOODS.NoMatch Then
   Continue = "N"
ElseIf TRIBESGOODS![ITEM_NUMBER] = 0 Then
   Continue = "N"
End If

Do Until Continue = "N"
   TRIBESGOODS.Edit
   If CONSTRUCTION_TYPE = "ENG" Then
      If TRIBESGOODS![ITEM_NUMBER] >= TOTAL_CLOTH Then
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TOTAL_CLOTH
         CLOTHUSED = CLOTHUSED + TOTAL_CLOTH
         If CLOTHUSED = TOTAL_CLOTH Then
            Continue = "N"
         ElseIf CLOTHUSED > TOTAL_CLOTH Then
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (CLOTHUSED - TOTAL_CLOTH)
            Continue = "N"
         End If
      Else
         Continue = "N"
      End If
   ElseIf WORKERS > 0 Then
      If TRIBESGOODS![ITEM_NUMBER] >= 2 Then
         If TOTAL_CLOTH - CLOTHUSED = 1 Then
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
         Else
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 2
         End If
         WORKERS = WORKERS - 1
         CLOTHUSED = CLOTHUSED + 2
         If CLOTHUSED = TOTAL_CLOTH Then
            Continue = "N"
         ElseIf CLOTHUSED > TOTAL_CLOTH Then
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (CLOTHUSED - TOTAL_CLOTH)
            Continue = "N"
         End If
      Else
         Continue = "N"
      End If
   Else
      Continue = "N"
   End If
   TRIBESGOODS.UPDATE
Loop

ERR_CALC_CLOTH_USED_CLOSE:
   Exit Function

ERR_CALC_CLOTH_USED:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_CLOTH_USED_CLOSE

End Function

Public Function CALC_DITCH()
On Error GoTo ERR_CALC_DITCH
TRIBE_STATUS = "CALC_DITCH"

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "DITCH"

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = "DITCH"
   HEXMAPCONST![1] = 0
   HEXMAPCONST.UPDATE
   HEXMAPCONST.index = "PRIMARYKEY"
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "DITCH"
   HEXMAPCONST.Edit
End If

HEXMAPCONST.Edit

Do Until WORKERS < 1
   If CONSTRUCTION = "DITCH" Then
      HEXMAPCONST![1] = HEXMAPCONST![1] + 1
      WORKERS = WORKERS - 1
   End If
Loop

HEXMAPCONST.UPDATE

ERR_CALC_DITCH_CLOSE:
   Exit Function

ERR_CALC_DITCH:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_DITCH_CLOSE

End Function

Function Calc_Engineering(CLAN, TRIBE, GTRIBE, ACTIVITY, ITEM, DISTINCTION, ACTIVES)
On Error GoTo ERR_ENGINEERING
TRIBE_STATUS = "Calc_Engineering"

Dim LOGS As Long
Dim STONES As Long
Dim COAL As Long
Dim BRASS As Long
Dim BRONZE As Long
Dim COPPER As Long
Dim IRON As Long
Dim LEAD As Long
Dim CLOTH As Long
Dim LEATHER As Long
Dim ROPES As Long
Dim LOGS_H As Long
Dim MILLSTONE As Long
Dim LOGS_Just_Used As Long
Dim STONES_Just_Used As Long
Dim COAL_Just_Used As Long
Dim BRASS_Just_Used As Long
Dim BRONZE_Just_Used As Long
Dim COPPER_Just_Used As Long
Dim IRON_Just_Used As Long
Dim LEAD_Just_Used As Long
Dim CLOTH_Just_Used As Long
Dim LEATHER_Just_Used As Long
Dim ROPES_Just_Used As Long
Dim LOGS_H_Just_Used As Long
Dim MILLSTONE_Just_Used As Long

Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst

Job = ACTIVITY
CONSTRUCTION = ITEM
BUILDING_TYPE = DISTINCTION
WORKERS = ACTIVES
TOTAL_WORKERS = 0
CONSTRUCTION_TYPE = "ENG"
NEW_CONSTRUCTION = "N"




CONSTCLAN = TOwning_Clan
CONSTTRIBE = TOwning_Tribe

' update turnactoutput
If Len(TurnActOutPut) > 20 Then
   TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people worked on " & StrConv(ITEM, vbProperCase)
Else
   TurnActOutPut = TurnActOutPut & " " & TActives & " effective people worked on " & StrConv(ITEM, vbProperCase)
End If

If Not CONSTTRIBE = TTRIBENUMBER Then
   TurnActOutPut = TurnActOutPut & " for tribe " & " using ("
Else
   TurnActOutPut = TurnActOutPut & " using ("
End If



WORKCLAN = CLAN
WORKTRIBE = TRIBE
GOODSTRIBE = GTRIBE

PARTS_FINISHED = 0
PARTS_TODO = 0
TOTAL_LOGS = 0
TOTAL_STONES = 0
TOTAL_COAL = 0
TOTAL_BRASS = 0
TOTAL_BRONZE = 0
TOTAL_COPPER = 0
TOTAL_IRON = 0
TOTAL_LEAD = 0
TOTAL_CLOTH = 0
TOTAL_LEATHER = 0
TOTAL_ROPES = 0
TOTAL_MILLSTONES = 0
LOGSUSED = 0
STONESUSED = 0
COALUSED = 0
BRASSUSED = 0
BRONZEUSED = 0
COPPERUSED = 0
IRONUSED = 0
LEADUSED = 0
CLOTHUSED = 0
LEATHERUSED = 0
ROPESUSED = 0
MILLSTONESUSED = 0
DONE = "N"
LOGS_Just_Used = 0
STONES_Just_Used = 0
COAL_Just_Used = 0
BRASS_Just_Used = 0
BRONZE_Just_Used = 0
COPPER_Just_Used = 0
IRON_Just_Used = 0
LEAD_Just_Used = 0
CLOTH_Just_Used = 0
LEATHER_Just_Used = 0
ROPES_Just_Used = 0
LOGS_H_Just_Used = 0
MILLSTONE_Just_Used = 0

'MSG = "CLAN = " & CLAN & " TRIBE = " & TRIBE & " GTRIBE = " & GTRIBE
'MsgBox (MSG)

'NEED TO ADD IN IMPLEMENTS...................
Call Process_Implement_Usage(TActivity, TItem, WORKERS, "NO")

If Mid(CONSTRUCTION, 5, 5) = "STONE" Then
   Call CALC_STONE_WALL
   DONE = "Y"
ElseIf Left(CONSTRUCTION, 5) = "DITCH" Then
   Call CALC_DITCH
   DONE = "Y"
ElseIf Right(CONSTRUCTION, 4) = "MOAT" Then
   Call CALC_MOAT
   DONE = "Y"
ElseIf Left(CONSTRUCTION, 11) = "PALISADE" Then
   Call CALC_WOOD_WALL
   DONE = "Y"
ElseIf Left(CONSTRUCTION, 9) = "MINESHAFT" Then
   Call Process_Mineshaft
   DONE = "Y"
End If

SEQ_NUMBER = 1
STOP_CONSTRUCTION = "N"

Set VALID_CONST = TVDB.OpenRecordset("VALID_BUILDINGS")
VALID_CONST.index = "PRIMARYKEY"
VALID_CONST.MoveFirst
VALID_CONST.Seek "=", CONSTRUCTION

If Not VALID_CONST.NoMatch Then
   SINGLE_CONSTRUCTION = VALID_CONST![One_Only]
End If

If DONE = "N" Then
Do Until STOP_CONSTRUCTION = "Y"
TOTAL_WORKERS = WORKERS

Do
   ConstructionTable.MoveFirst
   ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
   If SEQ_NUMBER > 1 Then
      If SINGLE_CONSTRUCTION = "Y" Then
         NEW_CONSTRUCTION = "N"
         ConstructionTable.MoveFirst
         ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, 1
         Exit Do
      End If
   End If
   If ConstructionTable.NoMatch Then
      If SEQ_NUMBER > 10 Then
         SEQ_NUMBER = 1
         ConstructionTable.AddNew
         ConstructionTable![TRIBE] = CONSTTRIBE
         ConstructionTable![CONSTRUCTION] = CONSTRUCTION
         ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
         ConstructionTable.UPDATE
         ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
         NEW_CONSTRUCTION = "Y"
         Exit Do
      Else
         ConstructionTable.AddNew
         ConstructionTable![TRIBE] = CONSTTRIBE
         ConstructionTable![CONSTRUCTION] = CONSTRUCTION
         ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
         ConstructionTable.UPDATE
         ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
         NEW_CONSTRUCTION = "Y"
         Exit Do
      End If
   Else
      NEW_CONSTRUCTION = "N"
      Exit Do
   End If
   SEQ_NUMBER = SEQ_NUMBER + 1
Loop
   
'MSG = "NEW CONSTRUCTION = " & NEW_CONSTRUCTION
'RESPONSE = MsgBox(MSG, True)

If NEW_CONSTRUCTION = "N" Then
   LOGSUSED = ConstructionTable![LOGS]
   STONESUSED = ConstructionTable![STONES]
   COALUSED = ConstructionTable![COAL]
   BRASSUSED = ConstructionTable![BRASS]
   BRONZEUSED = ConstructionTable![BRONZE]
   COPPERUSED = ConstructionTable![COPPER]
   IRONUSED = ConstructionTable![IRON]
   LEADUSED = ConstructionTable![LEAD]
   CLOTHUSED = ConstructionTable![CLOTH]
   LEATHERUSED = ConstructionTable![LEATHER]
   ROPESUSED = ConstructionTable![ROPES]
   MILLSTONESUSED = ConstructionTable![MILLSTONE]
Else
   LOGSUSED = 0
   STONESUSED = 0
   COALUSED = 0
   BRASSUSED = 0
   BRONZEUSED = 0
   COPPERUSED = 0
   IRONUSED = 0
   LEADUSED = 0
   CLOTHUSED = 0
   LEATHERUSED = 0
   ROPESUSED = 0
   MILLSTONESUSED = 0
End If

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE
TResearch = ActivitiesTable![research]

ItemsTable.index = "PRIMARYKEY"
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LOG"

If Not ItemsTable.NoMatch Then
   TOTAL_LOGS = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "STONE"
If Not ItemsTable.NoMatch Then
   TOTAL_STONES = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "COAL"
If Not ItemsTable.NoMatch Then
   TOTAL_COAL = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "BRASS"
If Not ItemsTable.NoMatch Then
   TOTAL_BRASS = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "BRONZE"
If Not ItemsTable.NoMatch Then
   TOTAL_BRONZE = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "COPPER"
If Not ItemsTable.NoMatch Then
   TOTAL_COPPER = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "IRON"
If Not ItemsTable.NoMatch Then
   TOTAL_IRON = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LEAD"
If Not ItemsTable.NoMatch Then
   TOTAL_LEAD = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "CLOTH"
If Not ItemsTable.NoMatch Then
   TOTAL_CLOTH = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LEATHER"
If Not ItemsTable.NoMatch Then
   TOTAL_LEATHER = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "ROPE"
If Not ItemsTable.NoMatch Then
   TOTAL_ROPES = ItemsTable![NUMBER]
End If
ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "MILLSTONE"
If Not ItemsTable.NoMatch Then
   TOTAL_MILLSTONES = ItemsTable![NUMBER]
End If

TRIBESGOODS.Seek "=", CONSTCLAN, GOODSTRIBE, "MINERAL", "BRASS"

If TOTAL_BRASS > 0 Then
   If TRIBESGOODS.NoMatch Then
      TOTAL_BRONZE = TOTAL_BRASS
      TOTAL_BRASS = 0
   ElseIf TRIBESGOODS![ITEM_NUMBER] <= 0 Then
      TOTAL_BRONZE = TOTAL_BRASS
      TOTAL_BRASS = 0
   End If
End If

If TOTAL_LOGS > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   If WORKERS > 0 Then
      Call CALC_ITEM_USED(LOGSUSED, "RAW", "LOG", TOTAL_LOGS, 2, "Y")
      ConstructionTable.Edit
      ConstructionTable![LOGS] = LOGSUSED
      ConstructionTable.UPDATE
      TurnActOutPut = TurnActOutPut & LOGSUSED & " logs, "
   End If
   If TOTAL_LOGS <= LOGSUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_STONES > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   If WORKERS >= 0 Then
      Call CALC_STONES_USED(STONESUSED)
      ConstructionTable.Edit
      ConstructionTable![STONES] = STONESUSED
      ConstructionTable.UPDATE
      TurnActOutPut = TurnActOutPut & STONESUSED & " stones, "
   End If
   If TOTAL_STONES <= STONESUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_COAL > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_ITEM_USED(COALUSED, "MINERAL", "COAL", TOTAL_COAL, TOTAL_COAL, "N")
   ConstructionTable.Edit
   ConstructionTable![COAL] = COALUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & COALUSED & " coal, "
   If TOTAL_COAL <= COALUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_BRASS > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_ITEM_USED(BRASSUSED, "MINERAL", "BRASS", TOTAL_BRASS, 10, "Y")
   ConstructionTable.Edit
   ConstructionTable![BRASS] = BRASSUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & BRASSUSED & " brass, "
   If TOTAL_BRASS <= BRASSUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_BRONZE > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(BRONZEUSED, "MINERAL", "BRONZE", TOTAL_BRONZE, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![BRONZE] = BRONZEUSED
      ConstructionTable.UPDATE
    TurnActOutPut = TurnActOutPut & BRONZEUSED & " bronze, "
   If TOTAL_BRONZE <= BRONZEUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_COPPER > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(COPPERUSED, "MINERAL", "COPPER", TOTAL_COPPER, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![COPPER] = COPPERUSED
      ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & COPPERUSED & " copper, "
   If TOTAL_COPPER <= COPPERUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_IRON > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(IRONUSED, "MINERAL", "IRON", TOTAL_IRON, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![IRON] = IRONUSED
      ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & IRONUSED & " iron, "
   If TOTAL_IRON <= IRONUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_LEAD > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(LEADUSED, "MINERAL", "LEAD", TOTAL_LEAD, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![LEAD] = LEADUSED
      ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & LEADUSED & " lead, "
   If TOTAL_LEAD <= LEADUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_CLOTH > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_CLOTH_USED(CLOTHUSED)
   ConstructionTable.Edit
   ConstructionTable![CLOTH] = CLOTHUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & CLOTHUSED & " cloth, "
   If TOTAL_CLOTH <= CLOTHUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_LEATHER > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_ITEM_USED(LEATHERUSED, "RAW", "LEATHER", TOTAL_LEATHER, TOTAL_LEATHER, "N")
   ConstructionTable.Edit
   ConstructionTable![LEATHER] = LEATHERUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & LEATHERUSED & " leather, "
   If TOTAL_LEATHER <= LEATHERUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_ROPES > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_ITEM_USED(ROPESUSED, "RAW", "ROPE", TOTAL_ROPES, TOTAL_ROPES, "N")
   ConstructionTable.Edit
   ConstructionTable![ROPES] = ROPESUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & ROPESUSED & " ropes, "
   If TOTAL_ROPES <= ROPESUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_MILLSTONES > 0 Then
   PARTS_TODO = PARTS_TODO + 1
   Call CALC_ITEM_USED(MILLSTONESUSED, "RAW", "MILLSTONE", TOTAL_MILLSTONES, TOTAL_MILLSTONES, "N")
   ConstructionTable.Edit
   ConstructionTable![MILLSTONE] = MILLSTONESUSED
   ConstructionTable.UPDATE
   TurnActOutPut = TurnActOutPut & MILLSTONESUSED & " millstones, "
   If TOTAL_MILLSTONES <= MILLSTONESUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If PARTS_TODO = PARTS_FINISHED Then
   x = CALC_BUILDING_ADDITION()
   ConstructionTable.Delete
   
   ' PERFORM RENUMBERING OF UNDERCONSTRUCTION TABLE
   ConstructionTable.MoveFirst

Do
   TRIBE = ConstructionTable![TRIBE]
   CONSTRUCTION = ConstructionTable![CONSTRUCTION]
   LOGS = ConstructionTable![LOGS]
   STONES = ConstructionTable![STONES]
   COAL = ConstructionTable![COAL]
   BRASS = ConstructionTable![BRASS]
   BRONZE = ConstructionTable![BRONZE]
   COPPER = ConstructionTable![COPPER]
   IRON = ConstructionTable![IRON]
   LEAD = ConstructionTable![LEAD]
   CLOTH = ConstructionTable![CLOTH]
   LEATHER = ConstructionTable![LEATHER]
   ROPES = ConstructionTable![ROPES]
   LOGS_H = ConstructionTable![LOG/H]
   MILLSTONE = ConstructionTable![MILLSTONE]
   ConstructionTable.Delete
      
   ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 1
   If ConstructionTable.NoMatch Then
      SEQ_NUMBER = 1
   Else
      ConstructionTable.MoveFirst
      ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 2
      If ConstructionTable.NoMatch Then
         SEQ_NUMBER = 2
      Else
         ConstructionTable.MoveFirst
         ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 3
         If ConstructionTable.NoMatch Then
            SEQ_NUMBER = 3
         Else
            ConstructionTable.MoveFirst
            ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 4
            If ConstructionTable.NoMatch Then
               SEQ_NUMBER = 4
            Else
               ConstructionTable.MoveFirst
               ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 5
               If ConstructionTable.NoMatch Then
                  SEQ_NUMBER = 5
               Else
                  ConstructionTable.MoveFirst
                  ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 6
                  If ConstructionTable.NoMatch Then
                     SEQ_NUMBER = 6
                  Else
                     ConstructionTable.MoveFirst
                     ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 7
                     If ConstructionTable.NoMatch Then
                        SEQ_NUMBER = 7
                     Else
                       ConstructionTable.MoveFirst
                       ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 8
                       If ConstructionTable.NoMatch Then
                          SEQ_NUMBER = 8
                       Else
                          ConstructionTable.MoveFirst
                          ConstructionTable.Seek "=", TRIBE, CONSTRUCTION, 9
                          If ConstructionTable.NoMatch Then
                             SEQ_NUMBER = 9
                          End If
                       End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
          
   ConstructionTable.AddNew
   ConstructionTable![TRIBE] = TRIBE
   ConstructionTable![CONSTRUCTION] = CONSTRUCTION
   ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
   ConstructionTable![LOGS] = LOGS
   ConstructionTable![STONES] = STONES
   ConstructionTable![COAL] = COAL
   ConstructionTable![BRASS] = BRASS
   ConstructionTable![BRONZE] = BRONZE
   ConstructionTable![COPPER] = COPPER
   ConstructionTable![IRON] = IRON
   ConstructionTable![LEAD] = LEAD
   ConstructionTable![CLOTH] = CLOTH
   ConstructionTable![LEATHER] = LEATHER
   ConstructionTable![ROPES] = ROPES
   ConstructionTable![LOG/H] = LOGS_H
   ConstructionTable![MILLSTONE] = MILLSTONE
   ConstructionTable.UPDATE
   ConstructionTable.MoveNext
   If ConstructionTable.EOF Then
      Exit Do
   End If
   
Loop

End If
Job = ACTIVITY
CONSTRUCTION = ITEM
BUILDING_TYPE = DISTINCTION
  
  If WORKERS = 1 Then
     WORKERS = 0
  End If
  
  If WORKERS <= 0 Then
     STOP_CONSTRUCTION = "Y"
  ElseIf WORKERS = TOTAL_WORKERS Then
     If WORKERS > 0 Then
        If SEQ_NUMBER = 2 Then
           STOP_CONSTRUCTION = "N"
        Else
           STOP_CONSTRUCTION = "Y"
       End If
     End If
     'MSG = "Still " & WORKERS & " workers left.  May be a problem"
     'MsgBox (MSG)
  ElseIf SINGLE_CONSTRUCTION = "Y" Then
     STOP_CONSTRUCTION = "Y"
  Else
     STOP_CONSTRUCTION = "N"
  End If
Loop

End If   ' FOR THE CHECK FOR DONE

VALID_CONST.Close

TurnActOutPut = TurnActOutPut & ") "

ERR_ENG_CLOSE:
   ' Ensure that HEXMAPCONST index is set back to Forthkey
   HEXMAPCONST.index = "FORTHKEY"
   Exit Function


ERR_ENGINEERING:
' 3167 - RECORD IS DELETED
If (Err = 3021) Or (Err = 3022) Or (Err = 3167) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_ENG_CLOSE
End If

End Function
Public Function CALC_MOAT()
On Error GoTo ERR_CALC_MOAT
TRIBE_STATUS = "CALC_MOAT"

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = CONSTRUCTION
   HEXMAPCONST![1] = 0
   HEXMAPCONST.UPDATE
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
End If

DITCH = 0
MOAT = 0
DOUBLE_MOAT = 0

HEXMAPCONST.Edit

If CONSTRUCTION = "MOAT" Then
   MOAT = HEXMAPCONST![1]
   DOUBLE_MOAT = 0
Else
   DOUBLE_MOAT = HEXMAPCONST![1]
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "MOAT"
   If Not HEXMAPCONST.NoMatch Then
       MOAT = HEXMAPCONST![1]
   End If
End If


HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "DITCH"

If Not HEXMAPCONST.NoMatch Then
   DITCH = HEXMAPCONST![1]
End If


HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
HEXMAPCONST.Edit

Do Until WORKERS < 1
   If DITCH = 0 Then
       If MOAT = 0 Then
           If CONSTRUCTION = "DOUBLE MOAT" Then
               If WORKERS >= 4 Then
                   HEXMAPCONST![1] = HEXMAPCONST![1] + 1
                   DOUBLE_MOAT = DOUBLE_MOAT + 1
                   WORKERS = WORKERS - 4
               Else
                   WORKERS = 0
               End If
           ElseIf CONSTRUCTION = "MOAT" Then
                  HEXMAPCONST![1] = HEXMAPCONST![1] + 1
                  MOAT = MOAT + 1
                 WORKERS = WORKERS - 2
           End If
       ElseIf CONSTRUCTION = "DOUBLE MOAT" Then
              If HEXMAPCONST![1] >= MOAT Then
                  If WORKERS >= 4 Then
                      HEXMAPCONST![1] = HEXMAPCONST![1] + 1
                      DOUBLE_MOAT = DOUBLE_MOAT + 1
                      WORKERS = WORKERS - 4
                  Else
                      WORKERS = 0
                  End If
              Else
                   HEXMAPCONST![1] = HEXMAPCONST![1] + 1
                    DOUBLE_MOAT = DOUBLE_MOAT + 1
                   WORKERS = WORKERS - 2
               End If
       ElseIf CONSTRUCTION = "MOAT" Then
              HEXMAPCONST![1] = HEXMAPCONST![1] + 1
              MOAT = MOAT + 1
             WORKERS = WORKERS - 2
       End If
   Else
      If HEXMAPCONST![1] >= DITCH Then
         If CONSTRUCTION = "MOAT" Then
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            MOAT = MOAT + 1
            WORKERS = WORKERS - 2
         End If
      Else
         If CONSTRUCTION = "MOAT" Then
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            MOAT = MOAT + 1
            WORKERS = WORKERS - 1
         End If
      End If
   End If

Loop
HEXMAPCONST.UPDATE

If DITCH > 0 Then
   If MOAT >= DITCH Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "DITCH"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

ERR_CALC_MOAT_CLOSE:
   Exit Function

ERR_CALC_MOAT:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_MOAT_CLOSE

End Function

Function CALC_SHIP_ADDITION()
On Error GoTo ERR_CALC_SHIP_ADDITION
TRIBE_STATUS = "CALC_SHIP_ADDITION"

Dim TINCREASE As Long

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", CONSTCLAN, GOODSTRIBE, "SHIP", CONSTRUCTION

Select Case CONSTRUCTION
Case "FISHER SS TO FISHER"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "FISHER", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "FISHER SS", "SUBTRACT", 1)
Case "H\FISHER SS TO H\FISHER"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\FISHER", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\FISHER SS", "SUBTRACT", 1)
Case "TRADER SS TO TRADER"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "TRADER", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "TRADER SS", "SUBTRACT", 1)
Case "H\TRADER SS TO H\TRADER"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\TRADER", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\TRADER SS", "SUBTRACT", 1)
Case "LONGSHIP SS TO LONGSHIP"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "LONGSHIP", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "LONGSHIP SS", "SUBTRACT", 1)
Case "LONGSHIP SS TO LONGSHIP"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\LONGSHIP", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\LONGSHIP SS", "SUBTRACT", 1)
Case "MERCHANT SS TO MERCHANT"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "MERCHANT", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "MERCHANT SS", "SUBTRACT", 1)
Case "H\MERCHANT SS TO H\MERCHANT"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\MERCHANT", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\MERCHANT SS", "SUBTRACT", 1)
Case "WARSHIP SS TO WARSHIP"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "MERCHANT", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "MERCHANT SS", "SUBTRACT", 1)
Case "H\WARSHIP SS TO H\WARSHIP"
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\MERCHANT", "ADD", 1)
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, "H\MERCHANT SS", "SUBTRACT", 1)
Case Else
   Call UPDATE_TRIBES_GOODS_TABLES(CONSTCLAN, GOODSTRIBE, CONSTRUCTION, "ADD", 1)
End Select

ERR_CALC_SHIP_ADDITION_CLOSE:
   Exit Function

ERR_CALC_SHIP_ADDITION:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_SHIP_ADDITION_CLOSE

End Function
Function Calc_Shipbuilding(CLAN, TRIBE, GTRIBE, ACTIVITY, ITEM, DISTINCTION, ACTIVES)
On Error GoTo ERR_SHIPBUILDING
TRIBE_STATUS = "Calc_Shipbuilding"

Job = ACTIVITY
CONSTRUCTION = ITEM
BUILDING_TYPE = DISTINCTION
WORKERS = ACTIVES
TOTAL_WORKERS = 0
CONSTRUCTION_TYPE = "SHIP"

CONSTCLAN = TOwning_Clan
CONSTTRIBE = TOwning_Tribe

WORKCLAN = CLAN
WORKTRIBE = TRIBE
GOODSTRIBE = GTRIBE

PARTS_FINISHED = 0
PARTS_TODO = 0
TOTAL_LOGS = 0
TOTAL_HLOGS = 0
TOTAL_STONES = 0
TOTAL_COAL = 0
TOTAL_BRASS = 0
TOTAL_BRONZE = 0
TOTAL_COPPER = 0
TOTAL_IRON = 0
TOTAL_LEAD = 0
TOTAL_CLOTH = 0
TOTAL_LEATHER = 0
TOTAL_ROPES = 0
LOGSUSED = 0
HLOGSUSED = 0
STONESUSED = 0
COALUSED = 0
BRASSUSED = 0
BRONZEUSED = 0
COPPERUSED = 0
IRONUSED = 0
LEADUSED = 0
CLOTHUSED = 0
LEATHERUSED = 0
ROPESUSED = 0
DONE = "N"

'MSG = "CLAN = " & CLAN & " TRIBE = " & TRIBE & " GTRIBE = " & GTRIBE
'MsgBox(MSG)

SEQ_NUMBER = 1
STOP_CONSTRUCTION = "N"
NEW_CONSTRUCTION = "N"

HEXMAPCONST.index = "FORTHKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, "SHIPYARD"

If HEXMAPCONST.NoMatch Then
   WORKERS = 0
   DONE = "Y"
ElseIf TActives > (HEXMAPCONST![1] * 10) Then
   ' UPDATE TURNOUTPUT
   If Len(TurnActOutPut) > 20 Then
      If Right(TurnActOutPut, 1) = " " Or Right(TurnActOutPut, 2) = ") " Or IsLetter(Right(TurnActOutPut, 1)) Then
         TurnActOutPut = TurnActOutPut & ", Too many workers allocated for the size of the shipyard "
         TurnActOutPut = TurnActOutPut & " while building " & StrConv(ITEM, vbProperCase)
      Else
         TurnActOutPut = TurnActOutPut & " Too many workers allocated for the size of the shipyard "
         TurnActOutPut = TurnActOutPut & " while building " & StrConv(ITEM, vbProperCase)
      End If
   Else
      TurnActOutPut = TurnActOutPut & " Too many workers allocated for the size of the shipyard "
      TurnActOutPut = TurnActOutPut & " while building " & StrConv(ITEM, vbProperCase)
   End If
   WORKERS = HEXMAPCONST![1] * 10
Else
   'UPDATE TURNOUTPUT
   If Len(TurnActOutPut) > 20 Then
      If Right(TurnActOutPut, 1) = " " Or Right(TurnActOutPut, 2) = ") " Or IsLetter(Right(TurnActOutPut, 1)) Then
         TurnActOutPut = TurnActOutPut & ", " & TActives & " effective people worked on " & StrConv(ITEM, vbProperCase)
      Else
         TurnActOutPut = TurnActOutPut & " " & TActives & " effective people worked on " & StrConv(ITEM, vbProperCase)
      End If
   Else
      TurnActOutPut = TurnActOutPut & " " & TActives & " effective people worked on " & StrConv(ITEM, vbProperCase)
   End If
End If

Call Process_Implement_Usage(TActivity, TItem, WORKERS, "NO")

HEXMAPCONST.index = "PRIMARYKEY"

ConstructionTable.MoveFirst

If DONE = "N" Then
Do Until STOP_CONSTRUCTION = "Y"
TOTAL_WORKERS = WORKERS

STOP_LOOP = "NO"
Do Until STOP_LOOP = "YES"
   ConstructionTable.MoveFirst
   ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
   If ConstructionTable.NoMatch Then
      If SEQ_NUMBER > 10 Then
         SEQ_NUMBER = 1
         ConstructionTable.AddNew
         ConstructionTable![TRIBE] = CONSTTRIBE
         ConstructionTable![CONSTRUCTION] = CONSTRUCTION
         ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
         ConstructionTable.UPDATE
         ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
         STOP_LOOP = "YES"
         NEW_CONSTRUCTION = "Y"
      Else
         ConstructionTable.AddNew
         ConstructionTable![TRIBE] = CONSTTRIBE
         ConstructionTable![CONSTRUCTION] = CONSTRUCTION
         ConstructionTable![SEQ NUMBER] = SEQ_NUMBER
         ConstructionTable.UPDATE
         ConstructionTable.Seek "=", CONSTTRIBE, CONSTRUCTION, SEQ_NUMBER
         STOP_LOOP = "YES"
         NEW_CONSTRUCTION = "Y"
      End If
   Else
      NEW_CONSTRUCTION = "N"
      STOP_LOOP = "YES"
   End If
   SEQ_NUMBER = SEQ_NUMBER + 1
Loop
   
If NEW_CONSTRUCTION = "N" Then
   HLOGSUSED = ConstructionTable![LOG/H]
   LOGSUSED = ConstructionTable![LOGS]
   STONESUSED = ConstructionTable![STONES]
   COALUSED = ConstructionTable![COAL]
   BRASSUSED = ConstructionTable![BRASS]
   BRONZEUSED = ConstructionTable![BRONZE]
   COPPERUSED = ConstructionTable![COPPER]
   IRONUSED = ConstructionTable![IRON]
   LEADUSED = ConstructionTable![LEAD]
   CLOTHUSED = ConstructionTable![CLOTH]
   LEATHERUSED = ConstructionTable![LEATHER]
   ROPESUSED = ConstructionTable![ROPES]
End If

ActivitiesTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE
TResearch = ActivitiesTable![research]

ItemsTable.index = "PRIMARYKEY"
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LOG"

If Not ItemsTable.NoMatch Then
   TOTAL_LOGS = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LOG/H"

If Not ItemsTable.NoMatch Then
   TOTAL_HLOGS = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "STONE"

If Not ItemsTable.NoMatch Then
   TOTAL_STONES = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "COAL"

If Not ItemsTable.NoMatch Then
   TOTAL_COAL = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "BRASS"

If Not ItemsTable.NoMatch Then
   TOTAL_BRASS = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "BRONZE"

If Not ItemsTable.NoMatch Then
   TOTAL_BRONZE = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "COPPER"

If Not ItemsTable.NoMatch Then
   TOTAL_COPPER = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "IRON"

If Not ItemsTable.NoMatch Then
   TOTAL_IRON = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LEAD"

If Not ItemsTable.NoMatch Then
   TOTAL_LEAD = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "CLOTH"

If Not ItemsTable.NoMatch Then
   TOTAL_CLOTH = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "LEATHER"

If Not ItemsTable.NoMatch Then
   TOTAL_LEATHER = ItemsTable![NUMBER]
End If

ItemsTable.MoveFirst
ItemsTable.Seek "=", Job, CONSTRUCTION, BUILDING_TYPE, "ROPE"

If Not ItemsTable.NoMatch Then
   TOTAL_ROPES = ItemsTable![NUMBER]
End If

TRIBESGOODS.Seek "=", CONSTCLAN, GOODSTRIBE, "MINERAL", "BRASS"

If TRIBESGOODS.NoMatch Then
   TOTAL_BRONZE = TOTAL_BRASS
   TOTAL_BRASS = 0
   BRONZEUSED = ConstructionTable![BRONZE] + ConstructionTable![BRASS]
ElseIf TRIBESGOODS![ITEM_NUMBER] <= 0 Then
   TOTAL_BRONZE = TOTAL_BRASS
   TOTAL_BRASS = 0
   BRONZEUSED = ConstructionTable![BRONZE] + ConstructionTable![BRASS]
End If

TRIBESGOODS.Seek "=", CONSTCLAN, GOODSTRIBE, "MINERAL", "LEAD"

If TRIBESGOODS.NoMatch Then
   TOTAL_COPPER = TOTAL_LEAD
   TOTAL_LEAD = 0
   COPPERUSED = ConstructionTable![COPPER] + ConstructionTable![LEAD]
ElseIf TRIBESGOODS![ITEM_NUMBER] <= 0 Then
   TOTAL_COPPER = TOTAL_LEAD
   TOTAL_LEAD = 0
   COPPERUSED = ConstructionTable![COPPER] + ConstructionTable![LEAD]
End If

If TOTAL_LOGS > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(LOGSUSED, "RAW", "LOG", TOTAL_LOGS, 2, "Y")
      ConstructionTable.Edit
      ConstructionTable![LOGS] = LOGSUSED
      ConstructionTable.UPDATE
   If TOTAL_LOGS <= LOGSUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_HLOGS > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(HLOGSUSED, "RAW", "LOG/H", TOTAL_HLOGS, 2, "Y")
      ConstructionTable.Edit
      ConstructionTable![H\LOGS] = HLOGSUSED
      ConstructionTable.UPDATE
   If TOTAL_HLOGS <= HLOGSUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If TOTAL_STONES > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_STONES_USED(STONESUSED)
      ConstructionTable.Edit
      ConstructionTable![STONES] = STONESUSED
      ConstructionTable.UPDATE
   If TOTAL_STONES <= STONESUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_COAL > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(COALUSED, "MINERAL", "COAL", TOTAL_COAL, TOTAL_COAL, "N")
      ConstructionTable.Edit
      ConstructionTable![COAL] = COALUSED
      ConstructionTable.UPDATE
   If TOTAL_COAL <= COALUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_BRASS > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(BRASSUSED, "MINERAL", "BRASS", TOTAL_BRASS, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![BRASS] = BRASSUSED
      ConstructionTable.UPDATE
   If TOTAL_BRASS <= BRASSUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_BRONZE > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(BRONZEUSED, "MINERAL", "BRONZE", TOTAL_BRONZE, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![BRONZE] = BRONZEUSED
      ConstructionTable.UPDATE
   If TOTAL_BRONZE <= BRONZEUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_COPPER > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(COPPERUSED, "MINERAL", "COPPER", TOTAL_COPPER, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![COPPER] = COPPERUSED
      ConstructionTable.UPDATE
   If TOTAL_COPPER <= COPPERUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_IRON > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(IRONUSED, "MINERAL", "IRON", TOTAL_IRON, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![IRON] = IRONUSED
      ConstructionTable.UPDATE
   If TOTAL_IRON <= IRONUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_LEAD > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(LEADUSED, "MINERAL", "LEAD", TOTAL_LEAD, 10, "Y")
      ConstructionTable.Edit
      ConstructionTable![LEAD] = LEADUSED
      ConstructionTable.UPDATE
   If TOTAL_LEAD <= LEADUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_CLOTH > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_CLOTH_USED(CLOTHUSED)
      ConstructionTable.Edit
      ConstructionTable![CLOTH] = CLOTHUSED
      ConstructionTable.UPDATE
   If TOTAL_CLOTH <= CLOTHUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_LEATHER > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(LEATHERUSED, "RAW", "LEATHER", TOTAL_LEATHER, TOTAL_LEATHER, "Y")
      ConstructionTable.Edit
      ConstructionTable![LEATHER] = LEATHERUSED
      ConstructionTable.UPDATE
   If TOTAL_LEATHER <= LEATHERUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If
If TOTAL_ROPES > 0 Then
   PARTS_TODO = PARTS_TODO + 1
      Call CALC_ITEM_USED(ROPESUSED, "RAW", "ROPE", TOTAL_ROPES, TOTAL_ROPES, "Y")
      ConstructionTable.Edit
      ConstructionTable![ROPES] = ROPESUSED
      ConstructionTable.UPDATE
   If TOTAL_ROPES <= ROPESUSED Then
      PARTS_FINISHED = PARTS_FINISHED + 1
   End If
End If

If PARTS_TODO = PARTS_FINISHED Then
   x = CALC_SHIP_ADDITION()
   ConstructionTable.Delete

End If
  If WORKERS <= 0 Then
     STOP_CONSTRUCTION = "Y"
  ElseIf WORKERS = TOTAL_WORKERS Then
     If SEQ_NUMBER = 2 Then
        STOP_CONSTRUCTION = "N"
     Else
        STOP_CONSTRUCTION = "Y"
     End If
  Else
        STOP_CONSTRUCTION = "N"
  End If
   LOGSUSED = 0
   STONESUSED = 0
   COALUSED = 0
   BRASSUSED = 0
   BRONZEUSED = 0
   COPPERUSED = 0
   IRONUSED = 0
   LEADUSED = 0
   CLOTHUSED = 0
   LEATHERUSED = 0
   ROPESUSED = 0

Loop
End If   ' FOR THE CHECK FOR DONE

ERR_SHIP_CLOSE:
   Exit Function


ERR_SHIPBUILDING:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume ERR_SHIP_CLOSE
End If


End Function
Public Function CALC_STONE_WALL()
On Error GoTo CALC_STONE_WALL_ERROR
Dim Stones_Used As Long
Dim Sand_Used As Long

TRIBE_STATUS = "CALC_STONE_WALL"

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst

CONCRETERS = 0
ENGINEERS = 0
Stones_Used = 0
Sand_Used = 0

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "CONCRETE")
        
If RESEARCH_FOUND = "Y" Then
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "SAND"
   TRIBESGOODS.Edit
   WALL_WORKERS = WORKERS
   Do Until WALL_WORKERS < 1
      If WALL_WORKERS >= 6 Then
         If TRIBESGOODS![ITEM_NUMBER] >= 10 Then
            CONCRETERS = CONCRETERS + 1
            ENGINEERS = ENGINEERS + 7.5
            WALL_WORKERS = WALL_WORKERS - 6
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 10
            TRIBESGOODS.UPDATE
            Sand_Used = Sand_Used + 10
         Else
            ENGINEERS = ENGINEERS + CONCRETERS + WALL_WORKERS
            WALL_WORKERS = 0
         End If
      Else
         ENGINEERS = ENGINEERS + CONCRETERS + WALL_WORKERS
         WALL_WORKERS = 0
      End If
   Loop

   WORKERS = CLng(ENGINEERS)
End If

HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.index = "MAP"
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", Tribes_Current_Hex
   If HEXMAPCONST.NoMatch Then
      HEXMAPCONST.AddNew
      HEXMAPCONST![MAP] = Tribes_Current_Hex
      HEXMAPCONST![CLAN] = CONSTCLAN
      HEXMAPCONST![TRIBE] = CONSTTRIBE
      HEXMAPCONST![CONSTRUCTION] = CONSTRUCTION
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
      HEXMAPCONST.index = "PRIMARYKEY"
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
      HEXMAPCONST.Edit
   ElseIf Not HEXMAPCONST![CLAN] = CONSTCLAN Then
      Do Until (HEXMAPCONST![CLAN] = CONSTCLAN) Or Not (HEXMAPCONST![MAP] = CURRENT_HEX)
         HEXMAPCONST.MoveNext
      Loop
      If Not HEXMAPCONST![CLAN] = CONSTCLAN Then
            HEXMAPCONST.AddNew
            HEXMAPCONST![MAP] = Tribes_Current_Hex
            HEXMAPCONST![CLAN] = CONSTCLAN
            HEXMAPCONST![TRIBE] = CONSTTRIBE
            HEXMAPCONST![CONSTRUCTION] = CONSTRUCTION
            HEXMAPCONST![1] = 0
            HEXMAPCONST.UPDATE
            HEXMAPCONST.index = "PRIMARYKEY"
            HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
            HEXMAPCONST.Edit
      End If
   End If
End If

LOG_WALL = 0
STONE10 = 0
STONE15 = 0
STONE20 = 0
STONE25 = 0
STONE30 = 0

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "PALISADE"

If HEXMAPCONST.NoMatch Then
   LOG_WALL = 0
Else
   LOG_WALL = HEXMAPCONST![1]
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "10' STONE WALL"

If HEXMAPCONST.NoMatch Then
   STONE10 = 0
Else
   STONE10 = HEXMAPCONST![1]
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "15' STONE WALL"

If HEXMAPCONST.NoMatch Then
   STONE15 = 0
Else
   STONE15 = HEXMAPCONST![1]
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "20' STONE WALL"

If HEXMAPCONST.NoMatch Then
   STONE20 = 0
Else
   STONE20 = HEXMAPCONST![1]
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "25' STONE WALL"

If HEXMAPCONST.NoMatch Then
   STONE25 = 0
Else
   STONE25 = HEXMAPCONST![1]
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "30' STONE WALL"

If HEXMAPCONST.NoMatch Then
   STONE30 = 0
Else
   STONE30 = HEXMAPCONST![1]
End If

If LOG_WALL > 0 Then
   If STONE10 >= LOG_WALL Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "PALISADE"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

If STONE10 > 0 Then
   If STONE15 >= STONE10 Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "10' STONE WALL"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

If STONE15 > 0 Then
   If STONE20 >= STONE15 Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "15' STONE WALL"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

If STONE20 > 0 Then
   If STONE25 >= STONE20 Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "20' STONE WALL"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

If STONE25 > 0 Then
   If STONE30 >= STONE25 Then
      HEXMAPCONST.MoveFirst
      HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "20' STONE WALL"
      HEXMAPCONST.Edit
      HEXMAPCONST![1] = 0
      HEXMAPCONST.UPDATE
   End If
End If

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = CONSTRUCTION
   HEXMAPCONST![1] = 0
   HEXMAPCONST.UPDATE
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
End If

HEXMAPCONST.Edit

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "STONE"
TRIBESGOODS.Edit

If CONSTRUCTION = "10' STONE WALL" Then
   Do Until WORKERS < 3
     If TRIBESGOODS![ITEM_NUMBER] >= 30 Then
         HEXMAPCONST.Edit
         HEXMAPCONST![1] = HEXMAPCONST![1] + 1
         HEXMAPCONST.UPDATE
         WORKERS = WORKERS - 3
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 30
         TRIBESGOODS.UPDATE
         Stones_Used = Stones_Used + 30
      Else
         WORKERS = 0
      End If
   Loop
ElseIf CONSTRUCTION = "15' STONE WALL" Then
   If STONE10 > 0 Then
      Do Until WORKERS < 6
         If TRIBESGOODS![ITEM_NUMBER] >= 45 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 6
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 45
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 45
         Else
            WORKERS = 0
         End If
      Loop
   Else
      Do Until WORKERS < 9
         If TRIBESGOODS![NUMBER] >= 75 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 9
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 75
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 75
         Else
            WORKERS = 0
         End If
      Loop
   End If
ElseIf CONSTRUCTION = "20' STONE WALL" Then
   If STONE15 > 0 Then
      Do Until WORKERS < 9
         If TRIBESGOODS![ITEM_NUMBER] >= 60 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 9
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 60
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 60
         Else
            WORKERS = 0
         End If
      Loop
   Else
      Do Until WORKERS < 18
         If TRIBESGOODS![ITEM_NUMBER] >= 135 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 18
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 135
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 135
         Else
            WORKERS = 0
         End If
      Loop
   End If
ElseIf CONSTRUCTION = "25' STONE WALL" Then
   If STONE20 > 0 Then
      Do Until WORKERS < 12
         If TRIBESGOODS![ITEM_NUMBER] >= 75 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 12
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 75
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 75
         Else
            WORKERS = 0
         End If
      Loop
   Else
      Do Until WORKERS < 24
         If TRIBESGOODS![ITEM_NUMBER] >= 150 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 24
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 150
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 150
         Else
            WORKERS = 0
         End If
      Loop
   End If
ElseIf CONSTRUCTION = "30' STONE WALL" Then
   If STONE25 > 0 Then
      Do Until WORKERS < 15
         If TRIBESGOODS![ITEM_NUMBER] >= 90 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 15
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 90
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 90
         Else
            WORKERS = 0
         End If
      Loop
   Else
      Do Until WORKERS < 30
         If TRIBESGOODS![ITEM_NUMBER] >= 150 Then
            HEXMAPCONST.Edit
            HEXMAPCONST![1] = HEXMAPCONST![1] + 1
            HEXMAPCONST.UPDATE
            WORKERS = WORKERS - 30
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 165
            TRIBESGOODS.UPDATE
            Stones_Used = Stones_Used + 165
         Else
            WORKERS = 0
         End If
      Loop
   End If
End If

TurnActOutPut = TurnActOutPut & Stones_Used & " stones"

CALC_STONE_WALL_ERROR_CLOSE:
   Exit Function


CALC_STONE_WALL_ERROR:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume CALC_STONE_WALL_ERROR_CLOSE
End If

End Function

Public Function CALC_STONES_USED(STONESUSED)
On Error GoTo ERR_CALC_STONES_USED
TRIBE_STATUS = "CALC_STONES_USED"

CONCRETERS = 0
ENGINEERS = 0

RESEARCH_FOUND = "N"

Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "CONCRETE")
        
If RESEARCH_FOUND = "Y" Then
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "SAND"
   TRIBESGOODS.Edit
   WALL_WORKERS = WORKERS
   Do Until WALL_WORKERS < 1
      If WALL_WORKERS >= 6 Then
         If TRIBESGOODS![ITEM_NUMBER] >= 10 Then
            CONCRETERS = CONCRETERS + 1
            ENGINEERS = ENGINEERS + 7.5
            WALL_WORKERS = WALL_WORKERS - 6
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 10
            TRIBESGOODS.UPDATE
         Else
            ENGINEERS = ENGINEERS + CONCRETERS + WALL_WORKERS
            WALL_WORKERS = 0
         End If
      Else
         ENGINEERS = ENGINEERS + CONCRETERS + WALL_WORKERS
         WALL_WORKERS = 0
      End If
   Loop

   WORKERS = CLng(ENGINEERS)
End If

If TOTAL_STONES > STONESUSED Then
   Continue = "Y"
Else
   Continue = "N"
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "STONE"
If TRIBESGOODS.NoMatch Then
   Msg = "Clan - " & WORKCLAN & " Tribe " & GOODSTRIBE & " does not have enough raw stone to build "
   Msg = Msg & TItem
   MsgBox (Msg)
Else
TRIBESGOODS.Edit

Do Until Continue = "N"
   If WORKERS > 0 Then
      If TRIBESGOODS![ITEM_NUMBER] >= 5 Then
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 5
         WORKERS = WORKERS - 1
         STONESUSED = STONESUSED + 5
         If STONESUSED = TOTAL_STONES Then
            Continue = "N"
         ElseIf STONESUSED > TOTAL_STONES Then
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (STONESUSED - TOTAL_STONES)
            Continue = "N"
         End If
      Else
         Continue = "N"
      End If
   Else
      Continue = "N"
   End If
Loop

TRIBESGOODS.UPDATE
End If

ERR_CALC_STONES_USED_CLOSE:
   Exit Function

ERR_CALC_STONES_USED:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_STONES_USED_CLOSE

End Function

Public Function CALC_WOOD_WALL()
On Error GoTo ERR_CALC_WOOD_WALL
Dim Wood_Used As Long

Wood_Used = 0

TRIBE_STATUS = "CALC_WOOD_WALL"

TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "LOG"
TRIBESGOODS.Edit

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = CONSTRUCTION
   HEXMAPCONST![1] = 0
   HEXMAPCONST.UPDATE
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, CONSTRUCTION
   HEXMAPCONST.Edit
End If

Do Until WORKERS < 1
   If CONSTRUCTION = "PALISADE" Then
      If TRIBESGOODS![ITEM_NUMBER] >= 3 Then
         HEXMAPCONST.Edit
         HEXMAPCONST![1] = HEXMAPCONST![1] + 1
         HEXMAPCONST.UPDATE
         WORKERS = WORKERS - 1
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 3
         TRIBESGOODS.UPDATE
         Wood_Used = Wood_Used + 3
      Else
         WORKERS = 0
      End If
   Else
      WORKERS = 0
   End If
Loop

TurnActOutPut = TurnActOutPut & Wood_Used & " logs, "

ERR_CALC_WOOD_WALL_CLOSE:
   Exit Function

ERR_CALC_WOOD_WALL:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_WOOD_WALL_CLOSE

End Function


Public Function Process_Mineshaft()
On Error GoTo ERR_Process_Mineshaft
TRIBE_STATUS = "Process_Mineshaft"

TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, "RAW", "LOG"
TRIBESGOODS.Edit

HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "MINESHAFT"

If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = Tribes_Current_Hex
   HEXMAPCONST![CLAN] = CONSTCLAN
   HEXMAPCONST![TRIBE] = CONSTTRIBE
   HEXMAPCONST![CONSTRUCTION] = "MINESHAFT"
   HEXMAPCONST![1] = 0
   HEXMAPCONST.UPDATE
   HEXMAPCONST.index = "PRIMARYKEY"
   HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "MINESHAFT"
   HEXMAPCONST.Edit
End If

HEXMAPCONST.Edit

Do Until WORKERS < 4
   If CONSTRUCTION = "MINESHAFT" Then
      If TRIBESGOODS![ITEM_NUMBER] >= 1 Then
         HEXMAPCONST.Edit
         HEXMAPCONST![1] = HEXMAPCONST![1] + 1
         HEXMAPCONST.UPDATE
         WORKERS = WORKERS - 4
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - 1
         TRIBESGOODS.UPDATE
      Else
         WORKERS = 0
      End If
   End If
Loop

ERR_Process_Mineshaft_CLOSE:
   Exit Function

ERR_Process_Mineshaft:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Mineshaft_CLOSE

End Function
Public Function CALC_ITEM_USED(AMT_ITEM_USED, TABLE, ITEM, TOTAL_ITEM, AMT_TO_USE, WORKERS_REQD)
On Error GoTo ERR_CALC_ITEM_USED
TRIBE_STATUS = "CALC_ITEM_USED"
'AMT_ITEM_USED is the amount of the item used in the construction
'TABLE is the table that the item resides in
'ITEM is the item being used
'TOTAL_ITEM is the total amount of that item to be used for the construction
'AMT_TO_USE is the amoount of theat item to use per worker.

TRIBESGOODS.Seek "=", WORKCLAN, GOODSTRIBE, TABLE, ITEM

If TOTAL_ITEM > AMT_ITEM_USED Then
   Continue = "Y"
Else
   Continue = "N"
End If

If TRIBESGOODS.NoMatch Then
   Continue = "N"
ElseIf TRIBESGOODS![ITEM_NUMBER] = 0 Then
   Continue = "N"
End If

Do Until Continue = "N"
   TRIBESGOODS.Edit
   If WORKERS > 0 Or WORKERS_REQD = "N" Then
      If TRIBESGOODS![ITEM_NUMBER] >= AMT_TO_USE Then
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - AMT_TO_USE
         If WORKERS_REQD = "Y" Then
            WORKERS = WORKERS - 1
         End If
         AMT_ITEM_USED = AMT_ITEM_USED + AMT_TO_USE
         If AMT_ITEM_USED = TOTAL_ITEM Then
            Continue = "N"
         ElseIf AMT_ITEM_USED >= TOTAL_ITEM Then
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + (AMT_ITEM_USED - TOTAL_ITEM)
            AMT_ITEM_USED = AMT_ITEM_USED + TOTAL_ITEM
            Continue = "N"
         End If
         TRIBESGOODS.UPDATE
      Else
         Continue = "N"
      End If
   Else
      Continue = "N"
   End If
Loop

ERR_CALC_ITEM_USED_CLOSE:
   Exit Function

ERR_CALC_ITEM_USED:
   Call A999_ERROR_HANDLING
   Resume ERR_CALC_ITEM_USED_CLOSE

End Function

Public Function Get_Research_Data(CLAN, TRIBE, research)
On Error GoTo ERR_Get_Research_Data
TRIBE_STATUS = "Get_Research_Data"

Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTAB.index = "PRIMARYKEY"
COMPRESTAB.MoveFirst

COMPRESTAB.Seek "=", TRIBE, research

RESEARCH_FOUND = "N"

If COMPRESTAB.NoMatch Then
   RESEARCH_FOUND = "N"
Else
   RESEARCH_FOUND = "Y"
End If

If IsNull(TRIBE) Then
   ' scroll clan for tribes
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", CLAN, CLAN

   Do While TRIBESINFO![CLAN] = TCLANNUMBER
      
      COMPRESTAB.MoveFirst
      COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], research
  
      If COMPRESTAB.NoMatch Then
         RESEARCH_FOUND = "N"
      Else
         RESEARCH_FOUND = "Y"
         Exit Do
      End If

      TRIBESINFO.MoveNext

    Loop

End If



ERR_Get_Research_Data_CLOSE:
   Exit Function

ERR_Get_Research_Data:
   Call A999_ERROR_HANDLING
   Resume ERR_Get_Research_Data_CLOSE

End Function

Public Function Get_Specialists_Info(CLAN, TRIBE, SPECIALIST)
On Error GoTo ERR_Get_Specialists_Info
TRIBE_STATUS = "Get_Specialists_Info"

NO_SPECIALISTS_FOUND = 0
BAKER_FOUND = 0
FARMER_FOUND = 0
FORESTER_FOUND = 0
HUNTER_FOUND = 0

Set TribesSpecialists = TVDBGM.OpenRecordset("TRIBES_SPECIALISTS")
TribesSpecialists.index = "PRIMARYKEY"
  If TribesSpecialists.BOF Then
      ' do nothing
  Else
      TribesSpecialists.MoveFirst
  End If
TribesSpecialists.Seek "=", CLAN, TRIBE, SPECIALIST

' check availability of specialist.
If TribesSpecialists.NoMatch Then
   SPECIALIST_FOUND = "N"
ElseIf SPECIALIST = "APIARIST" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "BAKER" Then
   SPECIALIST_FOUND = "Y"
   BAKER_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "BEEKEEPER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "BRICKLAYER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "DISTILLER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "FARMER" Then
   SPECIALIST_FOUND = "Y"
   FARMER_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "FORESTER" Then
   SPECIALIST_FOUND = "Y"
   FORESTER_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "FURRIER" Then
   SPECIALIST_FOUND = "Y"
   HUNTER_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "HERDER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "HUNTER" Then
   SPECIALIST_FOUND = "Y"
   HUNTER_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "METALWORKER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "MILLER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "MINER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "QUARRIER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "WEAVER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "WEAPONSMITH" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
ElseIf SPECIALIST = "WOODWORKER" Then
   SPECIALIST_FOUND = "Y"
   NO_SPECIALISTS_FOUND = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
Else
   SPECIALIST_FOUND = "Y"
   TSpecialists = TribesSpecialists![SPECIALISTS] - TribesSpecialists![SPECIALISTS_USED]
End If

ERR_Get_Specialists_Info_CLOSE:
   Exit Function

ERR_Get_Specialists_Info:
   Call A999_ERROR_HANDLING
   Resume ERR_Get_Specialists_Info_CLOSE

End Function

Public Function Perform_Cooking()
On Error GoTo ERR_Perform_Cooking
TRIBE_STATUS = "Perform_Cooking"

' NEED TO INSERT LIMITS ON COOKING.
' ONLY ALLOWED 10 COOKS PER LEVEL

Call PERFORM_COMMON("Y", "Y", "Y", 1, "NONE")

'if cooking baklava or napoleon then there should be a morale increase
' morale = morale + 0.01
If TItem = "BAKLAVA" Or TItem = "NAPOLEON" Then
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, TTRIBENUMBER
   TRIBESINFO.Edit
   TRIBESINFO![MORALE] = TRIBESINFO![MORALE] + 0.01
   TRIBESINFO.UPDATE
End If

ERR_Perform_Cooking_CLOSE:
   Exit Function

ERR_Perform_Cooking:
   Call A999_ERROR_HANDLING
   Resume ERR_Perform_Cooking_CLOSE

End Function

Public Function Process_People_Drinking()
On Error GoTo ERR_Process_People_Drinking
TRIBE_STATUS = "Process_People_Drinking"

Dim TERRAIN_FOUND As String
Dim Water_Required As Long
Dim Available_Water As Long

Available_Water = 0
Water_Required = 0

' at some stage we will have to identify if the tribe is in a state of siege.

' For drinking water we will need to identify the terrain
Select Case TRIBES_TERRAIN
Case "ARID"
   TERRAIN_FOUND = "Y"
Case "ARID HILLS"
   TERRAIN_FOUND = "Y"
Case "DESERT"
   TERRAIN_FOUND = "Y"
Case "OCEAN"
   TERRAIN_FOUND = "Y"
Case Else
   TERRAIN_FOUND = "N"
End Select

If TERRAIN_FOUND = "Y" Then
   ' the number of people and animals
   TRIBESINFO.Edit
   Water_Required = TRIBESINFO![WARRIORS] * 10
   Water_Required = Water_Required + (TRIBESINFO![ACTIVES] * 10)
   Water_Required = Water_Required + (TRIBESINFO![INACTIVES] * 10)
   Water_Required = Water_Required + (TRIBESINFO![SLAVE] * 5)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "GOAT") * 5)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "SHEEP") * 5)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "CATTLE") * 20)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "CAMEL") * 1)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "HORSE") * 20)
   Water_Required = Water_Required + (GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "ELEPHANT") * 40)
    
  
   POPULATION_STARVED = 0
   Available_Water = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, "WATER")

   MSG1 = "Drank "

   If Available_Water = 0 Then
      Call Check_Turn_Output("", " There is insufficient water to drink for this group this turn.", "", 0, "NO")
      Msg = GOODS_TRIBE & " IS MISSING " & Water_Required & " Water for people"
      MsgBox (Msg)
      MSG1 = MSG1 & "0 water, "
   ElseIf Water_Required > Available_Water Then
      Call Check_Turn_Output("", " There is insufficient water to drink for this group this turn.", "", 0, "NO")
      Msg = GOODS_TRIBE & "IS MISSING " & (Water_Required - Available_Water) & "Water for people to drink"
      MsgBox (Msg)
      MSG1 = MSG1 & Available_Water & " water, "
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WATER", "SUBTRACT", Available_Water)
   Else
      MSG1 = MSG1 & Water_Required & " water, "
      Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "WATER", "SUBTRACT", Water_Required)
   End If
End If

If TCLANNUMBER = "0330" Then
   Call Check_Turn_Output("", MSG1, "", 0, "NO")
End If

ERR_Process_People_Drinking_CLOSE:
   Exit Function

ERR_Process_People_Drinking:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_People_Drinking_CLOSE

End Function

Public Function Perform_Excavation()
On Error GoTo ERR_Perform_Excavation
TRIBE_STATUS = "Perform_Excavation"

Call PERFORM_COMMON("N", "Y", "Y", 1, "NONE")

ERR_Perform_Excavation_CLOSE:
   Exit Function

ERR_Perform_Excavation:
   Call A999_ERROR_HANDLING
   Resume ERR_Perform_Excavation_CLOSE


End Function

Public Function Perform_Conversions()
  TRIBE_STATUS = "Perform Conversion"
   If TItem = "Dogs" Then
      If TDistinction = "Herding Dogs" Then
         'check dogs
         ' check for research topic
         COMPRESTAB.MoveFirst
         COMPRESTAB.Seek "=", Skill_Tribe, "Herding Dogs"
         If Not COMPRESTAB.NoMatch Then
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "DOG"
            If TRIBESGOODS.NoMatch Then
               TDOGS = 0
            ElseIf TRIBESGOODS![ITEM_NUMBER] >= TActives Then
               TDOGS = TActives
            Else
               TDOGS = TRIBESGOODS![ITEM_NUMBER]
            End If
   
            'deduct dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Dog", "SUBTRACT", TDOGS)
            'add herding dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Herding Dog", "ADD", TDOGS)
               
            Call Check_Turn_Output(",", " Conversion of Dogs Performed ", "", 0, "YES")
         End If
      ElseIf TDistinction = "Hunting Dogs" Then
         ' check dogs and check for research topic
         COMPRESTAB.MoveFirst
         COMPRESTAB.Seek "=", Skill_Tribe, "Hunting Dogs"
         If Not COMPRESTAB.NoMatch Then
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "DOG"
            If TRIBESGOODS.NoMatch Then
               TDOGS = 0
            ElseIf TRIBESGOODS![ITEM_NUMBER] >= TActives Then
               TDOGS = TActives
            Else
               TDOGS = TRIBESGOODS![ITEM_NUMBER]
            End If
   
            'deduct dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Dog", "SUBTRACT", TDOGS)
            'add herding dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Hunting Dog", "ADD", TDOGS)
               
            Call Check_Turn_Output(",", " Conversion of Dogs Performed ", "", 0, "YES")
         End If
      ElseIf TDistinction = "Guard Dogs" Then
         ' check dogs and check for research topic
         COMPRESTAB.MoveFirst
         COMPRESTAB.Seek "=", Skill_Tribe, "Guard Dogs"
         If Not COMPRESTAB.NoMatch Then
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "ANIMAL", "DOG"
            If TRIBESGOODS.NoMatch Then
               TDOGS = 0
            ElseIf TRIBESGOODS![ITEM_NUMBER] >= TActives Then
               TDOGS = TActives
            Else
               TDOGS = TRIBESGOODS![ITEM_NUMBER]
            End If
   
            'deduct dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Dog", "SUBTRACT", TDOGS)
            'add herding dogs
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Guard Dog", "ADD", TDOGS)
               
            Call Check_Turn_Output(",", " Conversion of Dogs Performed ", "", 0, "YES")
         End If
      End If
ElseIf TItem = "Horse" Then
      ' check horse breeding last turn
      ' check for research topic
      COMPRESTAB.MoveFirst
      COMPRESTAB.Seek "=", Skill_Tribe, "Warhorse"
      If Not COMPRESTAB.NoMatch Then
          If Turn_Info_Req_NxTurn.BOF Then
              ' do nothing
          Else
              Turn_Info_Req_NxTurn.MoveFirst
          End If
          Turn_Info_Req_NxTurn.MoveFirst
          Turn_Info_Req_NxTurn.Seek "=", TCLANNUMBER, GOODS_TRIBE, "Horses Bred"
           
           If Turn_Info_Req_NxTurn.NoMatch Then
               THORSES = 0
           ElseIf Turn_Info_Req_NxTurn![ITEM_NUMBER] >= TActives Then
               THORSES = TActives
           Else
               THORSES = Turn_Info_Req_NxTurn![ITEM_NUMBER]
           End If
   
            'deduct horses
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Horse", "SUBTRACT", THORSES)
            'add warhorses
            Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "Warhorses", "ADD", THORSES)
               
            Call Check_Turn_Output(",", " Conversion of Horses to Warhorses ", "", 0, "YES")
         End If
End If
 



End Function

Public Function PERFORM_PACIFICATION(HEX_MAP, HEX_TERRAIN, MOVE_CLAN, MOVE_TRIBE)
Dim HEX_POPULATION As Integer

TRIBE_STATUS = "PERFORM PACIFICATION"
   
   Set TERRAINTABLE = TVDB.OpenRecordset("VALID_TERRAIN")
   TERRAINTABLE.index = "PRIMARYKEY"
   TERRAINTABLE.MoveFirst
   TERRAINTABLE.Seek "=", HEX_TERRAIN

   HEXMAPPOLITICS.MoveFirst
   HEXMAPPOLITICS.Seek "=", HEX_MAP
         
   If HEXMAPPOLITICS.NoMatch Then
      HEXMAPPOLITICS.AddNew
      HEXMAPPOLITICS![MAP] = HEX_MAP
      HEXMAPPOLITICS![PL_CLAN] = MOVE_CLAN
      HEXMAPPOLITICS![PL_TRIBE] = MOVE_TRIBE
      HEXMAPPOLITICS![PACIFICATION_LEVEL] = 0
      HEXMAPPOLITICS![POPULATION] = HEX_POPULATION
      HEXMAPPOLITICS.UPDATE
      HEXMAPPOLITICS.MoveFirst
      HEXMAPPOLITICS.Seek "=", HEX_MAP
      POLITICS_CLAN_CORRECT = "Y"
   ElseIf HEXMAPPOLITICS![PL_CLAN] = MOVE_CLAN And HEXMAPPOLITICS![PL_TRIBE] = MOVE_TRIBE Then
      POLITICS_CLAN_CORRECT = "Y"
   ElseIf HEXMAPPOLITICS![PL_CLAN] = "MOVE_CLAN" And HEXMAPPOLITICS![PL_TRIBE] = "MOVE_TRIBE" Then
      POLITICS_CLAN_CORRECT = "Y"
   ElseIf HEXMAPPOLITICS![PL_CLAN] = "N" Then
      POLITICS_CLAN_CORRECT = "Y"
   Else
      POLITICS_CLAN_CORRECT = "N"
   End If
    
   Call Get_Research_Data(TCLANNUMBER, Skill_Tribe, "Boat People")

   If RESEARCH_FOUND = "Y" Then
      If HEXMAPPOLITICS![POPULATION] = 0 Then
         HEX_POPULATION = 400
         HEXMAPPOLITICS.Edit
         HEXMAPPOLITICS![POPULATION] = HEX_POPULATION
         HEXMAPPOLITICS.UPDATE
      End If
   Else
      HEX_POPULATION = TERRAINTABLE![POPULATION]
   End If
  
   If POLITICS_CLAN_CORRECT = "Y" Then
      If Not IsNull(HEXMAPPOLITICS![PACIFICATION_LEVEL]) Then
         pllevel = HEXMAPPOLITICS![PACIFICATION_LEVEL]
      Else
         pllevel = 0
      End If
      If pllevel < 10 Then
         roll1 = DROLL(6, pllevel, 100, 0, DICE_TRIBE, 1, 0)
       
         If roll1 <= (110 - ((pllevel + 1) * 10)) Then
            If pllevel < 10 Then
               If pllevel = 0 Then
                  HEXMAPPOLITICS.Edit
                  HEXMAPPOLITICS![PL_CLAN] = MOVE_CLAN
                  HEXMAPPOLITICS![PL_TRIBE] = MOVE_TRIBE
                  HEXMAPPOLITICS![PACIFICATION_LEVEL] = pllevel + 1
                  HEXMAPPOLITICS![POPULATION] = HEX_POPULATION
                  HEXMAPPOLITICS.UPDATE
                  OutLine = OutLine & " Pacify Worked, "
               Else
                  HEXMAPPOLITICS.Edit
                  HEXMAPPOLITICS![PACIFICATION_LEVEL] = pllevel + 1
                  HEXMAPPOLITICS.UPDATE
                  OutLine = OutLine & " Pacify Worked, "
               End If
            End If
         Else
            OutLine = OutLine & " Pacify Failed, "
         End If
      Else
         OutLine = OutLine & " Pacify not reqd, "
      End If
   Else
      OutLine = OutLine & " Pacify Failed due to another Clan having pacified it, "
   End If
   

   
End Function
Public Function Initialise_Politics_Variables()
On Error GoTo ERR_Initialise_Politics_Variables
TRIBE_STATUS = "Initialise_Politics_Variables"

Ring1(1) = "N"
Ring1(2) = "NE"
Ring1(3) = "SE"
Ring1(4) = "S"
Ring1(5) = "SW"
Ring1(6) = "NW"

Ring2(1) = "N,N"
Ring2(2) = "N,NE"
Ring2(3) = "NE,NE"
Ring2(4) = "NE,SE"
Ring2(5) = "SE,SE"
Ring2(6) = "S,SE"
Ring2(7) = "S,S"
Ring2(8) = "S,SW"
Ring2(9) = "SW,SW"
Ring2(10) = "SW,NW"
Ring2(11) = "NW,NW"
Ring2(12) = "N,NW"

Ring3(1) = "N,N,N"
Ring3(2) = "N,N,NE"
Ring3(3) = "N,NE,NE"
Ring3(4) = "NE,NE,NE"
Ring3(5) = "NE,NE,SE"
Ring3(6) = "SE,SE,NE"
Ring3(7) = "SE,SE,SE"
Ring3(8) = "SE,SE,S"
Ring3(9) = "S,S,SE"
Ring3(10) = "S,S,S"
Ring3(11) = "S,S,SW"
Ring3(12) = "SW,SW,S"
Ring3(13) = "SW,SW,SW"
Ring3(14) = "SW,SW,NW"
Ring3(15) = "NW,NW,SW"
Ring3(16) = "NW,NW,NW"
Ring3(17) = "NW,NW,N"
Ring3(18) = "N,N,NW"

Ring4(1) = "N,N,N,N"
Ring4(2) = "N,N,N,NE"
Ring4(3) = "N,N,NE,NE"
Ring4(4) = "N,NE,NE,NE"
Ring4(5) = "NE,NE,NE,NE"
Ring4(6) = "NE,NE,NE,SE"
Ring4(7) = "NE,NE,SE,SE"
Ring4(8) = "NE,SE,SE,SE"
Ring4(9) = "SE,SE,SE,SE"
Ring4(10) = "SE,SE,SE,S"
Ring4(11) = "SE,SE,S,S"
Ring4(12) = "SE,S,S,S"
Ring4(13) = "S,S,S,S"
Ring4(14) = "SW,S,S,S"
Ring4(15) = "SW,SW,S,S"
Ring4(16) = "SW,SW,SW,S"
Ring4(17) = "SW,SW,SW,SW"
Ring4(18) = "NW,SW,SW,SW"
Ring4(19) = "NW,NW,SW,SW"
Ring4(20) = "NW,NW,NW,SW"
Ring4(21) = "NW,NW,NW,NW"
Ring4(22) = "NW,NW,NW,N"
Ring4(23) = "NW,NW,N,N"
Ring4(24) = "NW,N,N,N"


ERR_Initialise_Politics_Variables_CLOSE:
   Exit Function

ERR_Initialise_Politics_Variables:
   Call A999_ERROR_HANDLING
   Resume ERR_Initialise_Politics_Variables_CLOSE

End Function
Public Function Verify_Quantities()
On Error GoTo ERR_Verify_Quantities
TRIBE_STATUS = "Verify_Quantities"

'item to be used
'TGoods(Index1) = ItemsTable![GOOD]
'amount of item to be used
'TQuantity(Index1) = ItemsTable![NUMBER]
'number of items that can be made
'TNUMOCCURS

Index1 = 1
If TGoods(Index1) = "EMPTY" Then
   GoTo ERR_Verify_Quantities_CLOSE
End If
   
Do Until Index1 > 40
   
   'get item
   BRACKET = InStr(TGoods(Index1), "(")
   If BRACKET > 0 Then
      ITEM = Left(TGoods(Index1), (BRACKET - 1))
   Else
      ITEM = TGoods(Index1)
   End If
  
   'how many do we have?
   Num_Goods = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, GOODS_TRIBE, ITEM)
   
   If Num_Goods >= (TQuantity(Index1) * TNumItems) Then
      ' this is good
   ElseIf Num_Goods < (TQuantity(Index1) * TNumItems) Then
      TNUMOCCURS = Num_Goods / TQuantity(Index1)
      ActivesNeeded = TNUMOCCURS * TPeople
      NumItemsMade = TNUMOCCURS * TNumItems
   End If
   
   Index1 = Index1 + 1
   If TGoods(Index1) = "EMPTY" Then
      Exit Do
   End If
Loop

ERR_Verify_Quantities_CLOSE:
   Exit Function

ERR_Verify_Quantities:
   Call A999_ERROR_HANDLING
   Resume ERR_Verify_Quantities_CLOSE

End Function
Public Function Process_Clan_Research_Costs()
On Error GoTo ERR_Process_Clan_Research_Costs
TRIBE_STATUS = "Process_Clan_Research_Costs"
Dim Research_Tribe As String
Dim Research_Count As Integer
Dim Costs(8) As Long
Dim TOTALSILVER As Long

Costs(1) = 1
Costs(2) = 3
Costs(3) = 7
Costs(4) = 14
Costs(5) = 25
Costs(6) = 41
Costs(7) = 63
Costs(8) = 92

' Ok - for the clan, process the total research costs...

' TCLANNUMBER is the clan i am working with
' need to work through each tribes research within the clan
' Tribe checking has the clan and tribe info

Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
TRIBECHECK.index = "PRIMARYKEY"
TRIBECHECK.MoveFirst
TRIBECHECK.Seek "=", TCLANNUMBER, TCLANNUMBER

Set RESEARCHTABLE = TVDBGM.OpenRecordset("TRIBE_RESEARCH")
RESEARCHTABLE.index = "SECONDARYKEY"

'Research Costs are calculated according to this formula:
'200 * (1+2+4+7+11+16+22+29) x Silver.
'With the first research topic in each Tribe being free PLUS the second topic in the first 5 Tribes being free as well.
'Each number in brackets represents one topic beyond the free offerings (regardless of which sub-Tribe the topic is in).
'For example, 5 topics in the main Tribe is 3 topics beyond the free limits and would be 200 * (1+2+4) x Silver or 1400 Silver.
'Similarly 3 topics in the main Tribe and 3 topics in each of the first two sub-Tribes is 3 topics beyond the free limits and would be
'200 * (1+2+4) x Silver or 1400 Silver.  4 topics beyond the free limits would be 200 * (1+2+4+7) x Silver = 2800 Silver.

Research_Count = 0
TOTALSILVER = 0

Do Until TRIBECHECK![CLAN] <> TCLANNUMBER
   Research_Tribe = TRIBECHECK![TRIBE]
   count = 0
   RESEARCHTABLE.Seek "=", Research_Tribe
   If RESEARCHTABLE.NoMatch Then
      ' no research for tribe
      ' go to next tribe
   Else
      Do While RESEARCHTABLE![TRIBE] = Research_Tribe
         If Left(RESEARCHTABLE![TRIBE], 1) < 6 Then
            If RESEARCHTABLE![RESEARCH ATTEMPTED] = "Y" Then
               If count < 2 Then
                  count = count + 1
                  RESEARCHTABLE.MoveNext
               Else
                  ' count > 2
                  ' research must be costed
                  Research_Count = Research_Count + 1
                  count = count + 1
                  RESEARCHTABLE.MoveNext
               End If
            Else
               RESEARCHTABLE.MoveNext
            End If
         ElseIf RESEARCHTABLE![RESEARCH ATTEMPTED] = "Y" Then
            If count < 1 Then
               count = count + 1
               RESEARCHTABLE.MoveNext
            Else
               ' count > 1
               ' research must be costed
               Research_Count = Research_Count + 1
               count = count + 1
               RESEARCHTABLE.MoveNext
            End If
         Else
            RESEARCHTABLE.MoveNext
         End If
      Loop
   End If
   TRIBECHECK.MoveNext
Loop

' now we have the research_count, apply the calculation
If Research_Count > 8 Then
   Research_Count = 8
End If

TOTALSILVER = Costs(Research_Count) * 200
               
If TCLANNUMBER = "0330" Or TCLANNUMBER = "0445" Then
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"
   TRIBESGOODS.Edit
   MSG1 = ", " & TCLANNUMBER & "has " & TRIBESGOODS![ITEM_NUMBER] & "silver before research and has spent " & TOTALSILVER & " on Research this turn, "
   Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "ACTIVITIES", 2, MSG1, "Yes")
   TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - TOTALSILVER
   TRIBESGOODS.UPDATE
End If

ERR_Process_Clan_Research_Costs_CLOSE:
   RESEARCHTABLE.Close
   TRIBECHECK.Close
   Exit Function

ERR_Process_Clan_Research_Costs:
   Call A999_ERROR_HANDLING
   Resume ERR_Process_Clan_Research_Costs_CLOSE

End Function
' Checks the presence of Meeting House at any of adjacent hexes. Returns TRUE or FALSE (AlexD 24.06.24)
Public Function Look4AdjacentMH(sHex As String, sClan As String) As String
 Dim ADJ_HEXES(6) As String
 Dim i As Integer
 ' Get adjacent hexes
 ADJ_HEXES(1) = GET_MAP_NORTH(sHex)
 ADJ_HEXES(2) = GET_MAP_NORTH_EAST(sHex)
 ADJ_HEXES(3) = GET_MAP_SOUTH_EAST(sHex)
 ADJ_HEXES(4) = GET_MAP_SOUTH(sHex)
 ADJ_HEXES(5) = GET_MAP_SOUTH_WEST(sHex)
 ADJ_HEXES(6) = GET_MAP_NORTH_WEST(sHex)
 
 i = 1
 Do Until i > 6
        HEXMAPCONST.index = "FORTHKEY"
        HEXMAPCONST.MoveFirst
        HEXMAPCONST.Seek "=", ADJ_HEXES(i), sClan, "MEETING HOUSE"
        If Not HEXMAPCONST.NoMatch Then
                Look4AdjacentMH = "TRUE"
                Exit Function
        End If
        i = i + 1
 Loop
 Look4AdjacentMH = "FALSE"
End Function
' Checks  Terrain and looks for nearby village Returns TRUE or FALSE (AlexD 27.06.24)
Public Function CheckFarmingEligibility(sHex As String, sClan As String, sTribe As String) As String
 Dim sGoodsTribe As String
 
    CheckFarmingEligibility = "FALSE"
    'If Not ((TRIBES_TERRAIN = "PRAIRIE") Or (TRIBES_TERRAIN = "GRASSY_HILLS")) Then
    '    CheckFarmingEligibility = "FALSE"
    '    Exit Function
    'End If

' Check location for being village
    HEXMAPCONST.index = "FORTHKEY"
    HEXMAPCONST.MoveFirst
    HEXMAPCONST.Seek "=", sHex, sClan, "MEETING HOUSE"
    If Not HEXMAPCONST.NoMatch Then
        CheckFarmingEligibility = "TRUE"
        Exit Function
    End If

' Check for adjacent location MH
    If Look4AdjacentMH(sHex, sClan) Then
        CheckFarmingEligibility = "TRUE"
        Exit Function
    Else
        CheckFarmingEligibility = "FALSE" '" Can't plow without nearby village"
        Exit Function
    End If
' code below requires MH to belong to GT now it is disactivated by Exit Function
' Look for GT
    TRIBESINFO.MoveFirst
    TRIBESINFO.Seek "=", sClan, sTribe
    If IsNull(TRIBESINFO![GOODS TRIBE]) Then
        CheckFarmingEligibility = "FALSE" '" Can't plow without nearby village"
        Exit Function
    Else

    sGoodsTribe = TRIBESINFO![GOODS TRIBE]
    TRIBESINFO.MoveFirst
    TRIBESINFO.Seek "=", sClan, sGoodsTribe
    If TRIBESINFO.NoMatch Then 'GT not found
            CheckFarmingEligibility = "FALSE"
            Exit Function
    End If
    
 ' Checks that GT has MH
        
        
        
        HEXMAPCONST.index = "PRIMARYKEY"
        HEXMAPCONST.MoveFirst
        HEXMAPCONST.Seek "=", TRIBESINFO![CURRENT HEX], sClan, sGoodsTribe, "MEETING HOUSE"
        If Not HEXMAPCONST.NoMatch Then
            CheckFarmingEligibility = "TRUE"
            Exit Function
        Else
            CheckFarmingEligibility = "FALSE"
            Exit Function
        End If
    End If
End Function

' Checks Building for being "Container building" (AlexD 01.07.24)
Public Function isContainerBuilding(sBuilding As String) As String
    If sBuilding = "Refinery" Then
        isContainerBuilding = True
        Exit Function
    End If
    If sBuilding = "Charhouse" Then
        isContainerBuilding = True
        Exit Function
   End If
    If sBuilding = "Bakery" Then
        isContainerBuilding = True
        Exit Function
   End If
    If sBuilding = "Distillery" Then
        isContainerBuilding = True
        Exit Function
   End If
    If sBuilding = "Brickwork" Then
        isContainerBuilding = True
        Exit Function
   End If
   isContainerBuilding = False
End Function

' Checks Building for being "Installation Construction" (AlexD 04.07.24) Returns container name or FALSE
Public Function sContainerBuilding(sBuilding As String) As String
    If sBuilding = "BURNER" Then
        sContainerBuilding = "CHARHOUSE"
        Exit Function
    End If
    If sBuilding = "OVEN" Then
        sContainerBuilding = "Bakery"
        Exit Function
   End If
    ' If sBuilding = "STOVE" Then
    '    sContainerBuilding = "Bakery"
    '    Exit Function
    'End If
    If sBuilding = "STILL" Then
        sContainerBuilding = "Distillery"
        Exit Function
   End If
    If sBuilding = "KILN" Then
        sContainerBuilding = "Brickwork"
        Exit Function
   End If
    If sBuilding = "SMELTER" Then
        sContainerBuilding = "Refinery"
        Exit Function
   End If
   sContainerBuilding = "FALSE"
End Function


' Add Container Building (like Refinery) to HEXMAPCONST table.
' First building adds record to HEXMAPCONST table.
' All the next are marked by setting [1-10] slot values to 0 from -1
' AlexD 07.07.24
Public Function sAddContainerBuilding(sHex As String, sClan As String, sTribe As String, sBuilding As String) As String

Dim numSlotIndex As Long
sAddContainerBuilding = "DONE"

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", sHex, sClan, sTribe, sBuilding


If HEXMAPCONST.NoMatch Then
   HEXMAPCONST.AddNew
   HEXMAPCONST![MAP] = sHex
   HEXMAPCONST![CLAN] = sClan
   HEXMAPCONST![TRIBE] = sTribe
   HEXMAPCONST![CONSTRUCTION] = sBuilding
   numSlotIndex = 2 ' HEXMAPCONST[1] should be left 0 (that means structure is present)
   Do While numSlotIndex <= 10
        HEXMAPCONST(CStr(numSlotIndex)) = -1
        numSlotIndex = numSlotIndex + 1
   Loop
    HEXMAPCONST.UPDATE
    sAddContainerBuilding = "ADDED"
    Exit Function
Else
    HEXMAPCONST.Edit
   numSlotIndex = 1
   Do While numSlotIndex <= 10
        If HEXMAPCONST(CStr(numSlotIndex)) = -1 Then
            HEXMAPCONST(CStr(numSlotIndex)) = 0
            sAddContainerBuilding = "ADDED"
            HEXMAPCONST.UPDATE
            Exit Function
        End If
        numSlotIndex = numSlotIndex + 1
   Loop
sAddContainerBuilding = "MAX"
End If
HEXMAPCONST.UPDATE
End Function
' Add Installation Construction (like Smelter) to HEXMAPCONST table.
' Installation is added by advancing counter of slot  [1-10] at position defined by numPosition
' If specified slot is full (counter = 100) or contained doesn't exist (counter = -1)
' function vill try to find applicable slot
' AlexD 07.07.24
Public Function sAddInstallationConstruction(sHex As String, sClan As String, sTribe As String, sBuilding As String, numPosition As Long) As String
Dim numSlotIndex As Long
Dim sContainer As String
sAddInstallationConstruction = "DONE"
sContainer = sContainerBuilding(sBuilding)
If sContainer = "FALSE" Then
    sAddInstallationConstruction = "FALSE"
    Exit Function
End If
If numPosition = 0 Then numPosition = 1

HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", sHex, sClan, sTribe, sContainer
If HEXMAPCONST.NoMatch Then
    sAddInstallationConstruction = "FALSE"
    Exit Function
End If
HEXMAPCONST.Edit
If HEXMAPCONST(CStr(numPosition)) > -1 And HEXMAPCONST(CStr(numPosition)) < 100 Then
    HEXMAPCONST(CStr(numPosition)) = HEXMAPCONST(CStr(numPosition)) + 1
    sAddInstallationConstruction = "ADDED"
    HEXMAPCONST.UPDATE
    Exit Function
End If
    numSlotIndex = 1
   Do While numSlotIndex <= 10
    If HEXMAPCONST(CStr(numSlotIndex)) > -1 And HEXMAPCONST(CStr(numSlotIndex)) < 100 Then
        HEXMAPCONST(CStr(numSlotIndex)) = HEXMAPCONST(CStr(numSlotIndex)) + 1
        sAddInstallationConstruction = "ADDED"
        HEXMAPCONST.UPDATE
        Exit Function
    End If
        numSlotIndex = numSlotIndex + 1
   Loop

HEXMAPCONST.UPDATE
sAddInstallationConstruction = "MAX"
End Function

' Checks if it is possible to build tis building  (AlexD 04.07.24)
Public Function isCheckBuildingEligibility(sHex As String, sClan As String, sTribe As String, sBuilding As String) As String
Dim sContainer As String
Dim numSlotIndex As Long
' If this is defensive building you can build it anywhere
    If sBuilding = "10' STONE WALL" Or _
       sBuilding = "15' STONE WALL" Or _
       sBuilding = "20' STONE WALL" Or _
       sBuilding = "25' STONE WALL" Or _
       sBuilding = "30' STONE WALL" Or _
       sBuilding = "DITCH" Or _
       sBuilding = "MOAT" Or _
       sBuilding = "PALISADE" Or _
       sBuilding = "WOODEN TOWER" Or _
       sBuilding = "STONE TOWER" Or _
       sBuilding = "10' STONE WALL" Then
            isCheckBuildingEligibility = "True"
            Exit Function
        Else
            MsgBox "No match found!"
        End If

' if building is MH check there is none
    If sBuilding = "MEETING HOUSE" Then
        HEXMAPCONST.MoveFirst
        HEXMAPCONST.index = "TERTIARYKEY"
        HEXMAPCONST.Seek "=", sHex, "Meeting House"
        If HEXMAPCONST.NoMatch Then
            isCheckBuildingEligibility = "True"
            HEXMAPCONST.index = "PRIMARYKEY"
            Exit Function
        Else
            isCheckBuildingEligibility = " MH is already exist "
            HEXMAPCONST.index = "PRIMARYKEY"
            Exit Function
        End If
    End If
' Check that there is MH (not necessary belonging to your Clan)
    HEXMAPCONST.index = "TERTIARYKEY"
    HEXMAPCONST.MoveFirst
    HEXMAPCONST.Seek "=", sHex, "Meeting House"
    If HEXMAPCONST.NoMatch Then
        isCheckBuildingEligibility = " Can't build " & sBuilding & ": No MH "
        HEXMAPCONST.index = "PRIMARYKEY"
        Exit Function
    End If
' If building is a container you cant build more than 10 of them
    If isContainerBuilding(sBuilding) Then  'search for it
        HEXMAPCONST.index = "TERTIARYKEY"
        HEXMAPCONST.MoveFirst
        HEXMAPCONST.Seek "=", sHex, sBuilding
        If HEXMAPCONST.NoMatch Then
            isCheckBuildingEligibility = "True"
            HEXMAPCONST.index = "PRIMARYKEY"
            Exit Function
        End If
        If HEXMAPCONST![10] = -1 Then
            isCheckBuildingEligibility = "True"
            HEXMAPCONST.index = "PRIMARYKEY"
            Exit Function
        End If
        isCheckBuildingEligibility = " Maximum number of " & sBuilding & " already exist "
        Exit Function
    End If
    HEXMAPCONST.index = "PRIMARYKEY"
' if building is installation you can build it only if correspondent container building is exist
    sContainer = sContainerBuilding(sBuilding)
    If Not sContainer = "FALSE" Then
        HEXMAPCONST.index = "PRIMARYKEY"
        HEXMAPCONST.MoveFirst
        HEXMAPCONST.Seek "=", sHex, sClan, sTribe, sContainer
        If HEXMAPCONST.NoMatch Then
            isCheckBuildingEligibility = "No " & sContainer & " to build " & sBuilding
            HEXMAPCONST.index = "PRIMARYKEY"
            Exit Function
        End If

' And there are no more than 100 smelters in it
        numSlotIndex = 1
        Do While numSlotIndex <= 10
            If HEXMAPCONST(CStr(numSlotIndex)) < 100 Then
                isCheckBuildingEligibility = "True"
                Exit Function
            End If
            numSlotIndex = numSlotIndex + 1
        Loop
        isCheckBuildingEligibility = " Maximum number of " & sBuilding & " is reached "
        Exit Function
End If
isCheckBuildingEligibility = "True"
End Function

' Calculate Mouths of unit (without animals)
' AlexD 27.07.24
Public Function numCalcHumanMouths(sClan As String, sTribe As String) As Long
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", sClan, sTribe
   numCalcHumanMouths = 0
  
   If (TRIBESINFO!WARRIORS > 0) Then numCalcHumanMouths = numCalcHumanMouths + TRIBESINFO!WARRIORS
   If (TRIBESINFO!ACTIVES > 0) Then numCalcHumanMouths = numCalcHumanMouths + TRIBESINFO!ACTIVES
   If (TRIBESINFO!INACTIVES > 0) Then numCalcHumanMouths = numCalcHumanMouths + TRIBESINFO!INACTIVES
   If (TRIBESINFO!SLAVE > 0) Then numCalcHumanMouths = numCalcHumanMouths + TRIBESINFO!SLAVE
   If (TRIBESINFO!MERCENARIES > 0) Then numCalcHumanMouths = numCalcHumanMouths + TRIBESINFO!MERCENARIES
   
End Function

' Calculate total Mouths of GT of current unit (includes  animals and all dependent units)
' AlexD 27.07.24
Public Function numCalcTotalMouths(sClan As String, sTribe As String) As Long
Dim GOODSTRIBE As String
    numCalcTotalMouths = 0
    TRIBESINFO.MoveFirst
    TRIBESINFO.Seek "=", sClan, sTribe
    If TRIBESINFO.NoMatch Then
            Msg = Msg & "The Clan was " & sClan & " The Tribe was " & sTribe
            Msg = Msg & Chr(13) & Chr(10) & numCalcTotalMouths & " unit not found."
            MsgBox (Msg)
            Exit Function
    End If
   
   
    GOODSTRIBE = TRIBESINFO![GOODS TRIBE]
   ' Add all animals of GT
    TRIBESGOODS.MoveFirst
    TRIBESGOODS.Seek "=", sClan, GOODSTRIBE, "ANIMAL", "HERDING DOG"
    If Not TRIBESGOODS.NoMatch Then
        numCalcTotalMouths = numCalcTotalMouths + CLng(TRIBESGOODS![ITEM_NUMBER] / 2)
    End If
   
    TRIBESGOODS.MoveFirst
    TRIBESGOODS.Seek "=", sClan, GOODSTRIBE, "ANIMAL", "WARDOG"
    If Not TRIBESGOODS.NoMatch Then
        numCalcTotalMouths = numCalcTotalMouths + CLng(TRIBESGOODS![ITEM_NUMBER] / 2)
    End If

   ' find all units with this GT
   Do While TRIBESINFO![CLAN] = sClan ' All units of this clan
        If TRIBESINFO![GOODS TRIBE] = GOODSTRIBE Then
            numCalcTotalMouths = numCalcTotalMouths + numCalcHumanMouths(TRIBESINFO![CLAN], TRIBESINFO![TRIBE])
        End If
        TRIBESINFO.MoveNext
        If TRIBESINFO.EOF Then
            Exit Do
        End If
 
    Loop
End Function


' Calculate total amount of unstorable food of unit (Fish,Milk,Bread)
' AlexD 27.07.24
Public Function numCalcTotalUnstorableFood(sClan As String, sTribe As String) As Long
Dim GOODSTRIBE As String

    numCalcTotalUnstorableFood = 0
    TRIBESINFO.MoveFirst
    TRIBESINFO.Seek "=", sClan, sTribe
    If TRIBESINFO.NoMatch Then
            Msg = Msg & "The Clan was " & sClan & " The Tribe was " & sTribe
            Msg = Msg & Chr(13) & Chr(10) & numCalcTotalUnstorableFood & " unit not found."
            MsgBox (Msg)
            Exit Function
    End If
    GOODSTRIBE = TRIBESINFO![GOODS TRIBE]
   ' Add milk fich and bread of GT
    TRIBESGOODS.MoveFirst
    TRIBESGOODS.Seek "=", sClan, GOODSTRIBE, "RAW", "MILK"
    If Not TRIBESGOODS.NoMatch Then
        numCalcTotalUnstorableFood = numCalcTotalUnstorableFood + TRIBESGOODS![ITEM_NUMBER]
    End If
    TRIBESGOODS.MoveFirst
    TRIBESGOODS.Seek "=", sClan, GOODSTRIBE, "RAW", "FISH"
    If Not TRIBESGOODS.NoMatch Then
        numCalcTotalUnstorableFood = numCalcTotalUnstorableFood + TRIBESGOODS![ITEM_NUMBER]
    End If
    TRIBESGOODS.MoveFirst
    TRIBESGOODS.Seek "=", sClan, GOODSTRIBE, "FINISHED", "BREAD"
    If Not TRIBESGOODS.NoMatch Then
        numCalcTotalUnstorableFood = numCalcTotalUnstorableFood + TRIBESGOODS![ITEM_NUMBER]
    End If
  
End Function


' Calculates fish excess and salts it. Returns amount of  fish that remained unsalted
' AlexD 30.07.24
Public Function processSaltingExtraFish(sClan As String, sTribe As String, numFish As Long) As Long
Dim TOTAL_SALT As Long
Dim fishDemand As Long
Dim fishToSalt As Long
Dim maxFishToSalt As Long
Dim saltToUse As Long
' check for salt, if have salt then excess of fish may be salted.
TOTAL_SALT = GET_TRIBES_GOOD_QUANTITY(TCLANNUMBER, sTribe, "SALT")
fishDemand = numCalcTotalMouths(TCLANNUMBER, sTribe) - numCalcTotalUnstorableFood(TCLANNUMBER, sTribe)
If fishDemand < 0 Then
    fishDemand = 0
End If

'If TOTAL_SALT * 10 >= 1000 * SALTING_LEVEL Then
'    maxFishToSalt = 1000 * SALTING_LEVEL
'Else
    maxFishToSalt = TOTAL_SALT * 10
'End If


If TFishing >= fishDemand Then
    If TFishing - fishDemand <= maxFishToSalt Then ' salt it all
        fishToSalt = TFishing - fishDemand
    Else ' salt maxFishToSalt of fish
        fishToSalt = maxFishToSalt
    End If
        processSaltingExtraFish = TFishing - fishToSalt
        saltToUse = CLng(fishToSalt / 10)
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "SALT", "SUBTRACT", saltToUse)
'        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "FISH", "SUBTRACT", fishToSalt)
        Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, GOODS_TRIBE, "PROVS", "ADD", fishToSalt)
        Call Check_Turn_Output(" ", " and salted", " provs", fishToSalt, "NO")
        'TempOutput = TempOutput & ")"
        TurnActOutPut = TurnActOutPut & " (using " & fishToSalt & " fish and " & saltToUse & " salt)"
Else ' Nothing to salt
    processSaltingExtraFish = TFishing
End If
End Function
 
