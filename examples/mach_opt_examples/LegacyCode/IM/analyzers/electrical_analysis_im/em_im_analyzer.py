from time import time as clock_time
import os
from .FEMM_Solver import FEMM_Solver
import win32com.client
import logging

logger = logging.getLogger(__name__)
import numpy as np
# import population
from . import population


class IM_EM_Analysis():

    def __init__(self, configuration):
        self.configuration = configuration
        self.machine_variant = None
        self.operating_point = None

    def analyze(self, problem, counter=0):
        self.machine_variant = problem.machine
        self.operating_point = problem.operating_point
        problem.configuration = self.configuration
        ####################################################
        # 01 Setting project name and output folder
        ####################################################
        self.project_name = 'proj_%d_' % (counter)
        # Create output folder
        if not os.path.isdir(self.configuration['JMAG_csv_folder']):
            os.makedirs(self.configuration['JMAG_csv_folder'])

        self.machine_variant.fea_config_dict = self.configuration
        self.machine_variant.bool_initial_design = self.configuration['bool_initial_design']
        self.machine_variant.ID = self.project_name
        self.bool_run_in_JMAG_Script_Editor = False

        print('Run greedy_search_for_breakdown_slip...')
        femm_tic = clock_time()
        self.femm_solver = FEMM_Solver(self.machine_variant, flag_read_from_jmag=False, freq=500)  # eddy+static
        self.femm_solver.greedy_search_for_breakdown_slip(self.configuration['JMAG_csv_folder'], self.project_name,
                                                          bool_run_in_JMAG_Script_Editor=self.bool_run_in_JMAG_Script_Editor,
                                                          fraction=1)

        slip_freq_breakdown_torque, breakdown_torque, breakdown_force = self.femm_solver.wait_greedy_search(femm_tic)

        ## Setup JMAG Project

        app = win32com.client.Dispatch('designer.Application.181')
        if self.configuration['designer.Show'] == True:
            app.Show()
        else:
            app.Hide()

        self.app = app
        expected_project_file_path = self.project_name + 'JMAG'
        print(expected_project_file_path)
        if os.path.exists(expected_project_file_path):
            print(
                'JMAG project exists already. I learned my lessions. I will NOT delete it but create a new one with a different name instead.')
            # os.remove(expected_project_file_path)
            attempts = 2
            temp_path = expected_project_file_path[:-len('.jproj')] + 'attempts%d.jproj' % (attempts)
            while os.path.exists(temp_path):
                attempts += 1
                temp_path = expected_project_file_path[:-len('.jproj')] + 'attempts%d.jproj' % (attempts)

            expected_project_file_path = temp_path

        app.NewProject("Untitled")
        app.SaveAs(expected_project_file_path)

        DRAW_SUCCESS = self.draw_jmag_induction(app,
                                                counter,
                                                self.machine_variant,
                                                self.project_name)

        model = app.GetCurrentModel()
        tran2tss_study_name = self.project_name + 'Tran2TSS'
        study = self.add_TranFEAwi2TSS_study(self.machine_variant, 50.0, app, model, self.configuration['JMAG_csv_folder'],
                                                                                      tran2tss_study_name,
                                                                                      logger)
        app.SetCurrentStudy(tran2tss_study_name)
        study = app.GetCurrentStudy()
        self.mesh_study(self.machine_variant, app, model, study)

        if DRAW_SUCCESS == 0:
            # TODO: skip this model and its evaluation
            cost_function = 99999  # penalty
            logging.getLogger(__name__).warn('Draw Failed for %s-%s\nCost function penalty = %g.%s', self.project_name,
                                             self.project_name, cost_function, self.im_variant.show(toString=True))
            raise Exception(
                'Draw Failed: Are you working on the PC? Sometime you by mistake operate in the JMAG Geometry Editor, then it fails to draw.')
            return None
        elif DRAW_SUCCESS == -1:
            raise Exception(' DRAW_SUCCESS == -1:')

        # JMAG
        if app.NumModels() >= 1:
            model = app.GetModel(self.project_name)
        else:
            logger.error('there is no model yet for %s' % (self.project_name))
            raise Exception('why is there no model yet? %s' % (self.project_name))

        return slip_freq_breakdown_torque, breakdown_torque, breakdown_force


    def add_TranFEAwi2TSS_study(self, im_variant, slip_freq_breakdown_torque, app, model, dir_csv_output_folder, tran2tss_study_name, logger):

        # logger.debug('Slip frequency: %g = ' % (self.the_slip))
        self.the_slip = slip_freq_breakdown_torque / im_variant._machine_parameter_dict['DriveW_Freq']
        print("The Slip frequency is ", self.the_slip)
        # logger.debug('Slip frequency:    = %g???' % (self.the_slip))
        study_name = tran2tss_study_name

        model.CreateStudy("Transient2D", study_name)
        app.SetCurrentStudy(study_name)
        study = model.GetStudy(study_name)

        # SS-ATA
        study.GetStudyProperties().SetValue("ApproximateTransientAnalysis", 1) # psuedo steady state freq is for PWM drive to use
        study.GetStudyProperties().SetValue("SpecifySlip", 1)
        study.GetStudyProperties().SetValue("Slip", self.the_slip) # this will be overwritted later with "slip"
        study.GetStudyProperties().SetValue("OutputSteadyResultAs1stStep", 0)
        # study.GetStudyProperties().SetValue(u"TimePeriodicType", 2) # This is for TP-EEC but is not effective

        # misc
        study.GetStudyProperties().SetValue("ConversionType", 0)
        study.GetStudyProperties().SetValue("NonlinearMaxIteration", self.machine_variant.fea_config_dict["max_nonlinear_iteration"])
        study.GetStudyProperties().SetValue("ModelThickness", self.machine_variant._machine_parameter_dict["stack_length"]) # Stack Length

        # Material
        self.add_material(study)

        # Conditions - Motion
        self.the_speed = self.machine_variant._machine_parameter_dict["DriveW_Freq"]*60. / (0.5*self.machine_variant._machine_parameter_dict["DriveW_poles"]) * (1 - self.the_slip)
        study.CreateCondition("RotationMotion", "RotCon") # study.GetCondition(u"RotCon").SetXYZPoint(u"", 0, 0, 1) # megbox warning
        study.GetCondition("RotCon").SetValue("AngularVelocity", int(self.the_speed))
        study.GetCondition("RotCon").ClearParts()
        study.GetCondition("RotCon").AddSet(model.GetSetList().GetSet("Motion_Region"), 0)

        study.CreateCondition("Torque", "TorCon") # study.GetCondition(u"TorCon").SetXYZPoint(u"", 0, 0, 0) # megbox warning
        study.GetCondition("TorCon").SetValue("TargetType", 1)
        study.GetCondition("TorCon").SetLinkWithType("LinkedMotion", "RotCon")
        study.GetCondition("TorCon").ClearParts()

        study.CreateCondition("Force", "ForCon")
        study.GetCondition("ForCon").SetValue("TargetType", 1)
        study.GetCondition("ForCon").SetLinkWithType("LinkedMotion", "RotCon")
        study.GetCondition("ForCon").ClearParts()


        # Conditions - FEM Coils & Conductors (i.e. stator/rotor winding)
        self.add_circuit(app, model, study, bool_3PhaseCurrentSource=self.machine_variant.fea_config_dict["bool_3PhaseCurrentSource"])
        # quit()

        # True: no mesh or field results are needed
        study.GetStudyProperties().SetValue("OnlyTableResults", self.machine_variant.fea_config_dict['designer.OnlyTableResults'])

        # Linear Solver
        if False:
            # sometime nonlinear iteration is reported to fail and recommend to increase the accerlation rate of ICCG solver
            study.GetStudyProperties().SetValue("IccgAccel", 1.2)
            study.GetStudyProperties().SetValue("AutoAccel", 0)
        else:
            # this can be said to be super fast over ICCG solver.
            # https://www2.jmag-international.com/support/en/pdf/JMAG-Designer_Ver.17.1_ENv3.pdf
            study.GetStudyProperties().SetValue("DirectSolverType", 1) # require JMAG Designer version >17.05

        if self.fea_config_dict['designer.MultipleCPUs'] == True:
            # This SMP(shared memory process) is effective only if there are tons of elements. e.g., over 100,000.
            # too many threads will in turn make them compete with each other and slow down the solve. 2 is good enough for eddy current solve. 6~8 is enough for transient solve.
            study.GetStudyProperties().SetValue("UseMultiCPU", True)
            study.GetStudyProperties().SetValue("MultiCPU", 2)

        # # this is for the CAD parameters to rotate the rotor. the order matters for param_no to begin at 0.
        # if self.MODEL_ROTATE:
        #     self.add_cad_parameters(study)


        # 上一步的铁磁材料的状态作为下一步的初值，挺好，但是如果每一个转子的位置转过很大的话，反而会减慢非线性迭代。
        # 我们的情况是：0.33 sec 分成了32步，每步的时间大概在0.01秒，0.01秒乘以0.5*497 Hz = 2.485 revolution...
        # study.GetStudyProperties().SetValue(u"NonlinearSpeedup", 0) # JMAG17.1以后默认使用。现在后面密集的步长还多一点（32步），前面16步慢一点就慢一点呗！


        # two sections of different time step
        if True: # ECCE19
            number_of_steps_2ndTSS = self.fea_config_dict['designer.number_of_steps_2ndTSS']
            DM = app.GetDataManager()
            DM.CreatePointArray("point_array/timevsdivision", "SectionStepTable")
            refarray = [[0 for i in range(3)] for j in range(3)]
            refarray[0][0] = 0
            refarray[0][1] =    1
            refarray[0][2] =        50
            refarray[1][0] = 0.5/slip_freq_breakdown_torque #0.5 for 17.1.03l # 1 for 17.1.02y
            refarray[1][1] =    number_of_steps_2ndTSS                          # 16 for 17.1.03l #32 for 17.1.02y
            refarray[1][2] =        50
            refarray[2][0] = refarray[1][0] + 0.5/im_variant.DriveW_Freq #0.5 for 17.1.03l
            refarray[2][1] =    number_of_steps_2ndTSS  # also modify range_ss! # don't forget to modify below!
            refarray[2][2] =        50
            DM.GetDataSet("SectionStepTable").SetTable(refarray)
            number_of_total_steps = 1 + 2 * number_of_steps_2ndTSS # [Double Check] don't forget to modify here!
            study.GetStep().SetValue("Step", number_of_total_steps)
            study.GetStep().SetValue("StepType", 3)
            study.GetStep().SetTableProperty("Division", DM.GetDataSet("SectionStepTable"))

        else: # IEMDC19
            number_cycles_prolonged = 1 # 50
            DM = app.GetDataManager()
            DM.CreatePointArray("point_array/timevsdivision", "SectionStepTable")
            refarray = [[0 for i in range(3)] for j in range(4)]
            refarray[0][0] = 0
            refarray[0][1] =    1
            refarray[0][2] =        50
            refarray[1][0] = 1.0/slip_freq_breakdown_torque
            refarray[1][1] =    32
            refarray[1][2] =        50
            refarray[2][0] = refarray[1][0] + 1.0/im_variant.DriveW_Freq
            refarray[2][1] =    48 # don't forget to modify below!
            refarray[2][2] =        50
            refarray[3][0] = refarray[2][0] + number_cycles_prolonged/im_variant.DriveW_Freq # =50*0.002 sec = 0.1 sec is needed to converge to TranRef
            refarray[3][1] =    number_cycles_prolonged*self.fea_config_dict['designer.TranRef-StepPerCycle'] # =50*40, every 0.002 sec takes 40 steps
            refarray[3][2] =        50
            DM.GetDataSet("SectionStepTable").SetTable(refarray)
            study.GetStep().SetValue("Step", 1 + 32 + 48 + number_cycles_prolonged*self.fea_config_dict['designer.TranRef-StepPerCycle']) # [Double Check] don't forget to modify here!
            study.GetStep().SetValue("StepType", 3)
            study.GetStep().SetTableProperty("Division", DM.GetDataSet("SectionStepTable"))

        # add equations
        study.GetDesignTable().AddEquation("freq")
        study.GetDesignTable().AddEquation("slip")
        study.GetDesignTable().AddEquation("speed")
        study.GetDesignTable().GetEquation("freq").SetType(0)
        study.GetDesignTable().GetEquation("freq").SetExpression("%g"%((im_variant.DriveW_Freq)))
        study.GetDesignTable().GetEquation("freq").SetDescription("Excitation Frequency")
        study.GetDesignTable().GetEquation("slip").SetType(0)
        study.GetDesignTable().GetEquation("slip").SetExpression("%g"%(im_variant.the_slip))
        study.GetDesignTable().GetEquation("slip").SetDescription("Slip [1]")
        study.GetDesignTable().GetEquation("speed").SetType(1)
        study.GetDesignTable().GetEquation("speed").SetExpression("freq * (1 - slip) * %d"%(60/(im_variant.DriveW_poles/2)))
        study.GetDesignTable().GetEquation("speed").SetDescription("mechanical speed of four pole")

        # speed, freq, slip
        study.GetCondition("RotCon").SetValue("AngularVelocity", 'speed')
        if self.spec_input_dict['DPNV_or_SEPA']==False:
            app.ShowCircuitGrid(True)
            study.GetCircuit().GetComponent("CS%d"%(im_variant.DriveW_poles)).SetValue("Frequency", "freq")
            study.GetCircuit().GetComponent("CS%d"%(im_variant.BeariW_poles)).SetValue("Frequency", "freq")

        # max_nonlinear_iteration = 50
        # study.GetStudyProperties().SetValue(u"NonlinearMaxIteration", max_nonlinear_iteration)
        study.GetStudyProperties().SetValue("ApproximateTransientAnalysis", 1) # psuedo steady state freq is for PWM drive to use
        study.GetStudyProperties().SetValue("SpecifySlip", 1)
        study.GetStudyProperties().SetValue("OutputSteadyResultAs1stStep", 0)
        study.GetStudyProperties().SetValue("Slip", "slip") # overwrite with variables

        # # add other excitation frequencies other than 500 Hz as cases
        # for case_no, DriveW_Freq in enumerate([50.0, slip_freq_breakdown_torque]):
        #     slip = slip_freq_breakdown_torque / DriveW_Freq
        #     study.GetDesignTable().AddCase()
        #     study.GetDesignTable().SetValue(case_no+1, 0, DriveW_Freq)
        #     study.GetDesignTable().SetValue(case_no+1, 1, slip)

        # 你把Tran2TSS计算周期减半！
        # 也要在计算铁耗的时候选择1/4或1/2的数据！（建议1/4）
        # 然后，手动添加end step 和 start step，这样靠谱！2019-01-09：注意设置铁耗条件（iron loss condition）的Reference Start Step和End Step。

        # Iron Loss Calculation Condition
        # Stator
        if True:
            cond = study.CreateCondition("Ironloss", "IronLossConStator")
            cond.SetValue("RevolutionSpeed", "freq*60/%d"%(0.5*(im_variant.DriveW_poles)))
            cond.ClearParts()
            sel = cond.GetSelection()
            if bool_add_part_to_set_by_id == True:
                sel.SelectPart(self.id_statorCore)
            else:
                sel.SelectPartByPosition(-im_variant.Radius_OuterStatorYoke+EPS, 0 ,0)
            cond.AddSelected(sel)
            # Use FFT for hysteresis to be consistent with FEMM's results and to have a FFT plot
            cond.SetValue("HysteresisLossCalcType", 1)
            cond.SetValue("PresetType", 3) # 3:Custom
            # Specify the reference steps yourself because you don't really know what JMAG is doing behind you
            cond.SetValue("StartReferenceStep", number_of_total_steps+1-number_of_steps_2ndTSS*0.5) # 1/4 period <=> number_of_steps_2ndTSS*0.5
            cond.SetValue("EndReferenceStep", number_of_total_steps)
            cond.SetValue("UseStartReferenceStep", 1)
            cond.SetValue("UseEndReferenceStep", 1)
            cond.SetValue("Cyclicity", 4) # specify reference steps for 1/4 period and extend it to whole period
            cond.SetValue("UseFrequencyOrder", 1)
            cond.SetValue("FrequencyOrder", "1-50") # Harmonics up to 50th orders
        # Check CSV reults for iron loss (You cannot check this for Freq study) # CSV and save space
        study.GetStudyProperties().SetValue("CsvOutputPath", dir_csv_output_folder) # it's folder rather than file!
        study.GetStudyProperties().SetValue("CsvResultTypes", "Torque;Force;LineCurrent;TerminalVoltage;JouleLoss;TotalDisplacementAngle;JouleLoss_IronLoss;IronLoss_IronLoss;HysteresisLoss_IronLoss")
        study.GetStudyProperties().SetValue("DeleteResultFiles", self.fea_config_dict['delete_results_after_calculation'])
        # Terminal Voltage/Circuit Voltage: Check for outputing CSV results
        study.GetCircuit().CreateTerminalLabel("TerminalGroupBDU", 8, -13  + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        study.GetCircuit().CreateTerminalLabel("TerminalGroupBDV", 8, -11  + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        study.GetCircuit().CreateTerminalLabel("TerminalGroupBDW", 8, -9   + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        study.GetCircuit().CreateTerminalLabel("TerminalGroupACU", 23, -13 + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        study.GetCircuit().CreateTerminalLabel("TerminalGroupACV", 23, -11 + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        study.GetCircuit().CreateTerminalLabel("TerminalGroupACW", 23, -9  + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE)
        # Export Stator Core's field results only for iron loss calculation (the csv file of iron loss will be clean with this setting)
            # study.GetMaterial(u"Rotor Core").SetValue(u"OutputResult", 0) # at least one part on the rotor should be output or else a warning "the jplot file does not contains displacement results when you try to calc. iron loss on the moving part." will pop up, even though I don't add iron loss condition on the rotor.
        # study.GetMeshControl().SetValue(u"AirRegionOutputResult", 0)
        study.GetMaterial("Shaft").SetValue("OutputResult", 0)
        study.GetMaterial("Cage").SetValue("OutputResult", 0)
        study.GetMaterial("Coil").SetValue("OutputResult", 0)
        # Rotor
        if True:
            cond = study.CreateCondition("Ironloss", "IronLossConRotor")
            cond.SetValue("BasicFrequencyType", 2)
            cond.SetValue("BasicFrequency", "freq")
                # cond.SetValue(u"BasicFrequency", u"slip*freq") # this require the signal length to be at least 1/4 of slip period, that's too long!
            cond.ClearParts()
            sel = cond.GetSelection()
            if bool_add_part_to_set_by_id == True:
                sel.SelectPart(self.id_rotorCore)
            else:
                sel.SelectPartByPosition(-im_variant.Radius_Shaft-EPS, 0 ,0)
            cond.AddSelected(sel)
            # Use FFT for hysteresis to be consistent with FEMM's results
            cond.SetValue("HysteresisLossCalcType", 1)
            cond.SetValue("PresetType", 3)
            # Specify the reference steps yourself because you don't really know what JMAG is doing behind you
            cond.SetValue("StartReferenceStep", number_of_total_steps+1-number_of_steps_2ndTSS*0.5) # 1/4 period <=> number_of_steps_2ndTSS*0.5
            cond.SetValue("EndReferenceStep", number_of_total_steps)
            cond.SetValue("UseStartReferenceStep", 1)
            cond.SetValue("UseEndReferenceStep", 1)
            cond.SetValue("Cyclicity", 4) # specify reference steps for 1/4 period and extend it to whole period
            cond.SetValue("UseFrequencyOrder", 1)
            cond.SetValue("FrequencyOrder", "1-50") # Harmonics up to 50th orders
        self.study_name = study_name
        return study

    def add_material(self, study):
        if 'M19Gauge29' in self.machine_variant._materials_dict['stator_iron_mat']['core_material']:
            print('Inside stator iron material selection')
            study.SetMaterialByName("Stator Core", "M-19 Steel Gauge-29")
            study.GetMaterial("Stator Core").SetValue("Laminated", 1)
            study.GetMaterial("Stator Core").SetValue("LaminationFactor", 95)
                # study.GetMaterial(u"Stator Core").SetValue(u"UserConductivityValue", 1900000)
        if 'M19Gauge29' in self.machine_variant._materials_dict['rotor_iron_mat']['core_material']:

            study.SetMaterialByName("Rotor Core", "M-19 Steel Gauge-29")
            study.GetMaterial("Rotor Core").SetValue("Laminated", 1)
            study.GetMaterial("Rotor Core").SetValue("LaminationFactor", 95)

        # elif 'M15' in self.spec_input_dict['Steel']:
        #     study.SetMaterialByName("Stator Core", "M-15 Steel")
        #     study.GetMaterial("Stator Core").SetValue("Laminated", 1)
        #     study.GetMaterial("Stator Core").SetValue("LaminationFactor", 98)
        #
        #     study.SetMaterialByName("Rotor Core", "M-15 Steel")
        #     study.GetMaterial("Rotor Core").SetValue("Laminated", 1)
        #     study.GetMaterial("Rotor Core").SetValue("LaminationFactor", 98)
        #
        # elif self.spec_input_dict['Steel'] == 'Arnon5':
        #     study.SetMaterialByName("Stator Core", "Arnon5-final")
        #     study.GetMaterial("Stator Core").SetValue("Laminated", 1)
        #     study.GetMaterial("Stator Core").SetValue("LaminationFactor", 96)
        #
        #     study.SetMaterialByName("Rotor Core", "Arnon5-final")
        #     study.GetMaterial("Rotor Core").SetValue("Laminated", 1)
        #     study.GetMaterial("Rotor Core").SetValue("LaminationFactor", 96)

        else:
            msg = 'Warning: default material is used: DCMagnetic Type/50A1000.'
            print(msg)
            logging.getLogger(__name__).warn(msg)
            study.SetMaterialByName("Stator Core", "DCMagnetic Type/50A1000")
            study.GetMaterial("Stator Core").SetValue("UserConductivityType", 1)
            study.SetMaterialByName("Rotor Core", "DCMagnetic Type/50A1000")
            study.GetMaterial("Rotor Core").SetValue("UserConductivityType", 1)

        study.SetMaterialByName("Coil", "Copper")
        study.GetMaterial("Coil").SetValue("UserConductivityType", 1)

        study.SetMaterialByName("Cage", "Aluminium")
        study.GetMaterial("Cage").SetValue("EddyCurrentCalculation", 1)
        study.GetMaterial("Cage").SetValue("UserConductivityType", 1)
        study.GetMaterial("Cage").SetValue("UserConductivityValue", self.machine_variant._machine_parameter_dict['Bar_Conductivity'])

    def add_circuit(self, app, model, study, bool_3PhaseCurrentSource=True):
        # Circuit - Current Source
        app.ShowCircuitGrid(True)
        study.CreateCircuit()
        JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE = 0
        # 4 pole motor Qs=24 dpnv implemented by two layer winding (6 coils). In this case, drive winding has the same slot turns as bearing winding
        def circuit(Grouping, turns, Rs, ampD, ampB, freq, phase=0, CommutatingSequenceD=0, CommutatingSequenceB=0,
                    x=10, y=10 + JMAG_CIRCUIT_Y_POSITION_BIAS_FOR_CURRENT_SOURCE, bool_3PhaseCurrentSource=True):
            study.GetCircuit().CreateSubCircuit("Star Connection", "Star Connection %s" % (Grouping), x, y)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil1").SetValue("Turn",
                                                                                                               turns)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil1").SetValue(
                "Resistance", Rs)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil2").SetValue("Turn",
                                                                                                               turns)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil2").SetValue(
                "Resistance", Rs)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil3").SetValue("Turn",
                                                                                                               turns)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil3").SetValue(
                "Resistance", Rs)
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil1").SetName(
                "CircuitCoil%sU" % (Grouping))
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil2").SetName(
                "CircuitCoil%sV" % (Grouping))
            study.GetCircuit().GetSubCircuit("Star Connection %s" % (Grouping)).GetComponent("Coil3").SetName(
                "CircuitCoil%sW" % (Grouping))

            if bool_3PhaseCurrentSource == True:  # must use this for frequency analysis

                study.GetCircuit().CreateComponent("3PhaseCurrentSource", "CS%s" % (Grouping))
                study.GetCircuit().CreateInstance("CS%s" % (Grouping), x - 4, y + 1)
                study.GetCircuit().GetComponent("CS%s" % (Grouping)).SetValue("Amplitude", ampD + ampB)
                # study.GetCircuit().GetComponent("CS%s"%(Grouping)).SetValue("Frequency", "freq") # this is not needed for freq analysis # "freq" is a variable | 这个本来是可以用的，字符串"freq"的意思是用定义好的变量freq去代入，但是2020/07/07我重新搞Qs=36，p=3的Separate Winding的时候又不能正常工作的，circuit中设置的频率不是freq，而是0。
                study.GetCircuit().GetComponent("CS%s" % (Grouping)).SetValue("Frequency",
                                                                              self.DriveW_Freq)  # this is not needed for freq analysis # "freq" is a variable
                study.GetCircuit().GetComponent("CS%s" % (Grouping)).SetValue("PhaseU",
                                                                              phase)  # initial phase for phase U

                # Commutating sequence is essencial for the direction of the field to be consistent with speed: UVW rather than UWV
                # CommutatingSequenceD == 1 被我定义为相序UVW，相移为-120°，对应JMAG的"CommutatingSequence"为0，嗯，刚好要反一下，但我不改我的定义，因为DPNV那边（包括转子旋转方向）都已经按照这个定义测试好了，不改！
                # CommutatingSequenceD == 0 被我定义为相序UWV，相移为+120°，对应JMAG的"CommutatingSequence"为1，嗯，刚好要反一下，但我不改我的定义，因为DPNV那边（包括转子旋转方向）都已经按照这个定义测试好了，不改！
                if Grouping == 'Torque':
                    JMAGCommutatingSequence = 0
                    study.GetCircuit().GetComponent("CS%s" % (Grouping)).SetValue("CommutatingSequence",
                                                                                  JMAGCommutatingSequence)
                elif Grouping == 'Suspension':
                    JMAGCommutatingSequence = 1
                    study.GetCircuit().GetComponent("CS%s" % (Grouping)).SetValue("CommutatingSequence",
                                                                                  JMAGCommutatingSequence)
            else:
                I1 = "CS%s-1" % (Grouping)
                I2 = "CS%s-2" % (Grouping)
                I3 = "CS%s-3" % (Grouping)
                study.GetCircuit().CreateComponent("CurrentSource", I1)
                study.GetCircuit().CreateInstance(I1, x - 4, y + 3)
                study.GetCircuit().CreateComponent("CurrentSource", I2)
                study.GetCircuit().CreateInstance(I2, x - 4, y + 1)
                study.GetCircuit().CreateComponent("CurrentSource", I3)
                study.GetCircuit().CreateInstance(I3, x - 4, y - 1)

                phase_shift_drive = -120 if CommutatingSequenceD == 1 else 120
                phase_shift_beari = -120 if CommutatingSequenceB == 1 else 120

                func = app.FunctionFactory().Composite()
                f1 = app.FunctionFactory().Sin(ampD, freq,
                                               0 * phase_shift_drive)  # "freq" variable cannot be used here. So pay extra attension here when you create new case of a different freq.
                f2 = app.FunctionFactory().Sin(ampB, freq, 0 * phase_shift_beari)
                func.AddFunction(f1)
                func.AddFunction(f2)
                study.GetCircuit().GetComponent(I1).SetFunction(func)

                func = app.FunctionFactory().Composite()
                f1 = app.FunctionFactory().Sin(ampD, freq, 1 * phase_shift_drive)
                f2 = app.FunctionFactory().Sin(ampB, freq, 1 * phase_shift_beari)
                func.AddFunction(f1)
                func.AddFunction(f2)
                study.GetCircuit().GetComponent(I2).SetFunction(func)

                func = app.FunctionFactory().Composite()
                f1 = app.FunctionFactory().Sin(ampD, freq, 2 * phase_shift_drive)
                f2 = app.FunctionFactory().Sin(ampB, freq, 2 * phase_shift_beari)
                func.AddFunction(f1)
                func.AddFunction(f2)
                study.GetCircuit().GetComponent(I3).SetFunction(func)

            study.GetCircuit().CreateComponent("Ground", "Ground")
            study.GetCircuit().CreateInstance("Ground", x + 2, y + 1)

        # 这里电流幅值中的0.5因子源自DPNV导致的等于2的平行支路数。没有考虑到这一点，是否会对initial design的有效性产生影响？
        # 仔细看DPNV的接线，对于转矩逆变器，绕组的并联支路数为2，而对于悬浮逆变器，绕组的并联支路数为1。

        npb = self.machine_variant._machine_parameter_dict["number_parallel_branch"]
        nwl = self.machine_variant._machine_parameter_dict["number_winding_layer"]  # number of windign layers
        # if self.fea_config_dict['DPNV_separate_winding_implementation'] == True or self.spec_input_dict['DPNV_or_SEPA'] == False:
        if self.machine_variant._machine_parameter_dict['DPNV_or_SEPA'] == False:
            # either a separate winding or a DPNV winding implemented as a separate winding
            ampD = 0.5 * (self.DriveW_CurrentAmp / npb + self.BeariW_CurrentAmp)  # 为了代码能被四极电机和二极电机通用，代入看看就知道啦。
            ampB = -0.5 * (
                        self.DriveW_CurrentAmp / npb - self.BeariW_CurrentAmp)  # 关于符号，注意下面的DriveW对应的circuit调用时的ampB前还有个负号！
            if bool_3PhaseCurrentSource != True:
                raise Exception('Logic Error Detected.')
        else:
            '[B]: DriveW_CurrentAmp is set.'
            # case: DPNV as an actual two layer winding
            ampD = self.machine_variant._machine_parameter_dict["DriveW_CurrentAmp"] / npb
            ampB = self.machine_variant._machine_parameter_dict["BeariW_CurrentAmp"]
            if bool_3PhaseCurrentSource != False:
                raise Exception('Logic Error Detected.')

        Function = 'GroupAC' if self.machine_variant._machine_parameter_dict['DPNV_or_SEPA'] == True else 'Torque'
        circuit(Function, self.machine_variant._machine_parameter_dict["DriveW_zQ"] / nwl, bool_3PhaseCurrentSource=bool_3PhaseCurrentSource,
                Rs=self.machine_variant._machine_parameter_dict["DriveW_Rs"], ampD=ampD,
                ampB=-ampB, freq=self.machine_variant._machine_parameter_dict["DriveW_Freq"], phase=0,
                CommutatingSequenceD=self.machine_variant._machine_parameter_dict["CommutatingSequenceD"],
                CommutatingSequenceB=self.machine_variant._machine_parameter_dict["CommutatingSequenceB"])
        Function = 'GroupBD' if self.machine_variant._machine_parameter_dict['DPNV_or_SEPA'] == True else 'Suspension'
        circuit(Function, self.machine_variant._machine_parameter_dict["BeariW_turns"] / nwl, bool_3PhaseCurrentSource=bool_3PhaseCurrentSource,
                Rs=self.machine_variant._machine_parameter_dict["BeariW_Rs"], ampD=ampD,
                ampB=+ampB, freq=self.machine_variant._machine_parameter_dict["BeariW_Freq"], phase=0,
                CommutatingSequenceD=self.machine_variant._machine_parameter_dict["CommutatingSequenceD"],
                CommutatingSequenceB=self.machine_variant._machine_parameter_dict["CommutatingSequenceB"],
                x=25)  # CS4 corresponds to uauc (conflict with following codes but it does not matter.)

        # Link FEM Coils to Coil Set
        # if self.fea_config_dict['DPNV_separate_winding_implementation'] == True or self.spec_input_dict['DPNV_or_SEPA'] == False:
        if self.machine_variant._machine_parameter_dict['DPNV_or_SEPA'] == False:  # Separate Winding
            def link_FEMCoils_2_CoilSet(Function, l1, l2):
                # link between FEM Coil Condition and Circuit FEM Coil
                for UVW in ['U', 'V', 'W']:
                    which_phase = "Cond-%s-%s" % (Function, UVW)
                    study.CreateCondition("FEMCoil", which_phase)

                    condition = study.GetCondition(which_phase)
                    condition.SetLink("CircuitCoil%s%s" % (Function, UVW))
                    condition.GetSubCondition("untitled").SetName("Coil Set 1")
                    condition.GetSubCondition("Coil Set 1").SetName("delete")
                count = 0
                dict_dir = {'+': 1, '-': 0, 'o': None}
                # select the part to assign the FEM Coil condition
                for UVW, UpDown in zip(l1, l2):
                    count += 1
                    if dict_dir[UpDown] is None:
                        # print 'Skip', UVW, UpDown
                        continue
                    which_phase = "Cond-%s-%s" % (Function, UVW)
                    condition = study.GetCondition(which_phase)
                    condition.CreateSubCondition("FEMCoilData", "%s-Coil Set %d" % (Function, count))
                    subcondition = condition.GetSubCondition("%s-Coil Set %d" % (Function, count))
                    subcondition.ClearParts()
                    layer_symbol = 'LX' if Function == 'Torque' else 'LY'
                    subcondition.AddSet(model.GetSetList().GetSet(f"Coil{layer_symbol}{UVW}{UpDown} {count}"), 0)
                    subcondition.SetValue("Direction2D", dict_dir[UpDown])
                # clean up
                for UVW in ['U', 'V', 'W']:
                    which_phase = "Cond-%s-%s" % (Function, UVW)
                    condition = study.GetCondition(which_phase)
                    condition.RemoveSubCondition("delete")

            link_FEMCoils_2_CoilSet('Torque',
                                    self.dict_coil_connection['layer X phases'],
                                    self.dict_coil_connection['layer X signs'])
            link_FEMCoils_2_CoilSet('Suspension',
                                    self.dict_coil_connection['layer Y phases'],
                                    self.dict_coil_connection['layer Y signs'])
        else:  # DPNV Winding
            # 两个改变，一个是激励大小的改变（本来是200A 和 5A，现在是205A和195A），
            # 另一个绕组分组的改变，现在的A相是上层加下层为一相，以前是用俩单层绕组等效的。

            # Link FEM Coils to Coil Set as double layer short pitched winding
            # Create FEM Coil Condition
            # here we map circuit component `Coil2A' to FEM Coil Condition 'phaseAuauc
            # here we map circuit component `Coil4A' to FEM Coil Condition 'phaseAubud
            for suffix in ['GroupAC', 'GroupBD']:  # 仍然需要考虑poles，是因为为Coil设置Set那里的代码还没有更新。这里的2和4等价于leftlayer和rightlayer。
                for UVW in ['U', 'V', 'W']:
                    study.CreateCondition("FEMCoil", 'phase' + UVW + suffix)
                    # link between FEM Coil Condition and Circuit FEM Coil
                    condition = study.GetCondition('phase' + UVW + suffix)
                    condition.SetLink("CircuitCoil%s%s" % (suffix, UVW))
                    condition.GetSubCondition("untitled").SetName("delete")

            count = 0  # count indicates which slot the current rightlayer is in.
            index = 0
            dict_dir = {'+': 1, '-': 0}
            coil_pitch = self.machine_variant._machine_parameter_dict['pitch']  # self.dict_coil_connection[0]
            # select the part (via `Set') to assign the FEM Coil condition
            for UVW, UpDown in zip(self.machine_variant._machine_parameter_dict["layer_phases"], self.machine_variant._machine_parameter_dict["layer_polarity"]):

                count += 1
                print("UVW is", UVW)
                if self.machine_variant._machine_parameter_dict["coil_groups"][index] == 1:
                    suffix = 'GroupAC'
                else:
                    suffix = 'GroupBD'
                condition = study.GetCondition('phase' + UVW + suffix)

                # right layer
                condition.CreateSubCondition("FEMCoilData", "Coil Set Layer X %d" % (count))
                subcondition = condition.GetSubCondition("Coil Set Layer X %d" % (count))
                subcondition.ClearParts()
                subcondition.AddSet(model.GetSetList().GetSet("CoilLX%s%s %d" % (UVW, UpDown, count)),
                                    0)  # poles=4 means right layer, rather than actual poles
                subcondition.SetValue("Direction2D", dict_dir[UpDown])

                # left layer
                if coil_pitch <= 0:
                    raise Exception('把永磁电机circuit部分的代码移植过来！')
                if count + coil_pitch <= self.Qs:
                    count_leftlayer = count + coil_pitch
                    index_leftlayer = index + coil_pitch
                else:
                    count_leftlayer = int(count + coil_pitch - self.Qs)
                    index_leftlayer = int(index + coil_pitch - self.Qs)
                # 右层导体的电流方向是正，那么与其串联的一个coil_pitch之处的左层导体就是负！不需要再检查l_leftlayer2了~
                if UpDown == '+':
                    UpDown = '-'
                else:
                    UpDown = '+'
                condition.CreateSubCondition("FEMCoilData", "Coil Set Layer Y %d" % (count_leftlayer))
                subcondition = condition.GetSubCondition("Coil Set Layer Y %d" % (count_leftlayer))
                subcondition.ClearParts()
                subcondition.AddSet(model.GetSetList().GetSet("CoilLY%s%s %d" % (UVW, UpDown, count_leftlayer)),
                                    0)  # poles=2 means left layer, rather than actual poles
                subcondition.SetValue("Direction2D", dict_dir[UpDown])
                index += 1

                # double check
                if self.wily.layer_Y_phases[index_leftlayer] != UVW:
                    raise Exception('Bug in winding diagram.')
            # clean up
            for suffix in ['GroupAC', 'GroupBD']:
                for UVW in ['U', 'V', 'W']:
                    condition = study.GetCondition('phase' + UVW + suffix)
                    condition.RemoveSubCondition("delete")
            # raise Exception('Test DPNV PE.')

        # Condition - Conductor (i.e. rotor winding)
        for ind in range(int(self.Qr)):
            natural_ind = ind + 1
            study.CreateCondition("FEMConductor", "CdctCon %d" % (natural_ind))
            study.GetCondition("CdctCon %d" % (natural_ind)).GetSubCondition("untitled").SetName("Conductor Set 1")
            study.GetCondition("CdctCon %d" % (natural_ind)).GetSubCondition("Conductor Set 1").ClearParts()
            study.GetCondition("CdctCon %d" % (natural_ind)).GetSubCondition("Conductor Set 1").AddSet(
                model.GetSetList().GetSet("Bar %d" % (natural_ind)), 0)

        # Condition - Conductor - Grouping
        study.CreateCondition("GroupFEMConductor", "CdctCon_Group")
        for ind in range(int(self.Qr)):
            natural_ind = ind + 1
            study.GetCondition("CdctCon_Group").AddSubCondition("CdctCon %d" % (natural_ind), ind)

        # Link Conductors to Circuit
        def place_conductor(x, y, name):
            study.GetCircuit().CreateComponent("FEMConductor", name)
            study.GetCircuit().CreateInstance(name, x, y)

        def place_resistor(x, y, name, end_ring_resistance):
            study.GetCircuit().CreateComponent("Resistor", name)
            study.GetCircuit().CreateInstance(name, x, y)
            study.GetCircuit().GetComponent(name).SetValue("Resistance", end_ring_resistance)

        rotor_phase_name_list = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        X = 40;
        Y = 60;
        #
        if self.spec_input_dict[
            'PoleSpecificNeutral'] == True:  # Our proposed pole-specific winding with a neutral plate

            wily_Qr = winding_layout.pole_specific_winding_with_neutral(self.Qr, self.DriveW_poles / 2,
                                                                        self.BeariW_poles / 2,
                                                                        self.spec_input_dict['coil_pitch_y_Qr'])
            for ind, pair in enumerate(wily_Qr.pairs):
                X += -12
                # Y += -12
                if len(pair) != 2:
                    # double layer pole-specific winding with neutral plate

                    # print(f'This double layer rotor winding is goting to be implemented as its single layer equivalent (a coil with {len(pair)} conductors) with the neutral plate structure.\n')

                    for index, k in enumerate(pair):
                        string_conductor = f"Conductor{rotor_phase_name_list[ind]}{index + 1}"
                        place_conductor(X, Y - 9 * index, string_conductor)
                        study.GetCondition("CdctCon %d" % (k)).SetLink(string_conductor)

                        # no end ring resistors to behave like FEMM model
                        study.GetCircuit().CreateInstance("Ground", X - 5, Y - 2)
                        study.GetCircuit().CreateWire(X - 2, Y, X - 5, Y)
                        study.GetCircuit().CreateWire(X - 5, Y, X - 2, Y - 9 * index)
                        study.GetCircuit().CreateWire(X + 2, Y, X + 2, Y - 9 * index)

                    # quit()
                else:
                    # single layer pole-specific winding with neutral plate
                    i, j = pair
                    place_conductor(X, Y, "Conductor%s1" % (rotor_phase_name_list[ind]))
                    place_conductor(X, Y - 9, "Conductor%s2" % (rotor_phase_name_list[ind]))
                    study.GetCondition("CdctCon %d" % (i)).SetLink("Conductor%s1" % (rotor_phase_name_list[ind]))
                    study.GetCondition("CdctCon %d" % (j)).SetLink("Conductor%s2" % (rotor_phase_name_list[ind]))

                    # no end ring resistors to behave like FEMM model
                    study.GetCircuit().CreateWire(X + 2, Y, X + 2, Y - 9)
                    study.GetCircuit().CreateInstance("Ground", X - 5, Y - 2)
                    study.GetCircuit().CreateWire(X - 2, Y, X - 5, Y)
                    study.GetCircuit().CreateWire(X - 5, Y, X - 2, Y - 9)

            if self.End_Ring_Resistance != 0:  # setting a small value to End_Ring_Resistance is a bad idea (slow down the solver). Instead, don't model it
                raise Exception('With end ring is not implemented.')
        else:
            # 下边的方法只适用于Qr是p的整数倍的情况，比如Qr=28，p=3就会出错哦。
            if self.spec_input_dict['PS_or_SC'] == True:  # Chiba's conventional pole-specific winding
                if self.DriveW_poles == 2:
                    for i in range(int(self.no_slot_per_pole)):
                        Y += -12
                        place_conductor(X, Y, "Conductor%s1" % (rotor_phase_name_list[i]))
                        # place_conductor(X, Y-3, u"Conductor%s2"%(rotor_phase_name_list[i]))
                        # place_conductor(X, Y-6, u"Conductor%s3"%(rotor_phase_name_list[i]))
                        place_conductor(X, Y - 9, "Conductor%s2" % (rotor_phase_name_list[i]))

                        if self.End_Ring_Resistance == 0:  # setting a small value to End_Ring_Resistance is a bad idea (slow down the solver). Instead, don't model it
                            # no end ring resistors to behave like FEMM model
                            study.GetCircuit().CreateWire(X + 2, Y, X + 2, Y - 9)
                            # study.GetCircuit().CreateWire(X-2, Y-3, X-2, Y-6)
                            # study.GetCircuit().CreateWire(X+2, Y-6, X+2, Y-9)
                            study.GetCircuit().CreateInstance("Ground", X - 5, Y - 2)
                            study.GetCircuit().CreateWire(X - 2, Y, X - 5, Y)
                            study.GetCircuit().CreateWire(X - 5, Y, X - 2, Y - 9)
                        else:
                            raise Exception('With end ring is not implemented.')
                elif self.DriveW_poles == 4:  # poles = 4
                    for i in range(int(self.no_slot_per_pole)):
                        Y += -12
                        place_conductor(X, Y, "Conductor%s1" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 3, "Conductor%s2" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 6, "Conductor%s3" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 9, "Conductor%s4" % (rotor_phase_name_list[i]))

                        if self.End_Ring_Resistance == 0:  # setting a small value to End_Ring_Resistance is a bad idea (slow down the solver). Instead, don't model it
                            # no end ring resistors to behave like FEMM model
                            study.GetCircuit().CreateWire(X + 2, Y, X + 2, Y - 3)
                            study.GetCircuit().CreateWire(X - 2, Y - 3, X - 2, Y - 6)
                            study.GetCircuit().CreateWire(X + 2, Y - 6, X + 2, Y - 9)
                            study.GetCircuit().CreateInstance("Ground", X - 5, Y - 2)
                            study.GetCircuit().CreateWire(X - 2, Y, X - 5, Y)
                            study.GetCircuit().CreateWire(X - 5, Y, X - 2, Y - 9)
                        else:
                            place_resistor(X + 4, Y, "R_%s1" % (rotor_phase_name_list[i]), self.End_Ring_Resistance)
                            place_resistor(X - 4, Y - 3, "R_%s2" % (rotor_phase_name_list[i]), self.End_Ring_Resistance)
                            place_resistor(X + 4, Y - 6, "R_%s3" % (rotor_phase_name_list[i]), self.End_Ring_Resistance)
                            place_resistor(X - 4, Y - 9, "R_%s4" % (rotor_phase_name_list[i]), self.End_Ring_Resistance)

                            study.GetCircuit().CreateWire(X + 6, Y, X + 2, Y - 3)
                            study.GetCircuit().CreateWire(X - 6, Y - 3, X - 2, Y - 6)
                            study.GetCircuit().CreateWire(X + 6, Y - 6, X + 2, Y - 9)
                            study.GetCircuit().CreateWire(X - 6, Y - 9, X - 7, Y - 9)
                            study.GetCircuit().CreateWire(X - 2, Y, X - 7, Y)
                            study.GetCircuit().CreateInstance("Ground", X - 7, Y - 2)
                            # study.GetCircuit().GetInstance(u"Ground", ini_ground_no+i).RotateTo(90)
                            study.GetCircuit().CreateWire(X - 7, Y, X - 6, Y - 9)

                elif self.DriveW_poles == 6:  # poles = 6
                    for i in range(int(self.no_slot_per_pole)):
                        # Y += -3*(self.DriveW_poles-1) # work for poles below 6
                        X += 10  # tested for End_Ring_Resistance==0 only
                        place_conductor(X, Y, "Conductor%s1" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 3, "Conductor%s2" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 6, "Conductor%s3" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 9, "Conductor%s4" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 12, "Conductor%s5" % (rotor_phase_name_list[i]))
                        place_conductor(X, Y - 15, "Conductor%s6" % (rotor_phase_name_list[i]))

                        if self.End_Ring_Resistance == 0:  # setting a small value to End_Ring_Resistance is a bad idea (slow down the solver). Instead, don't model it
                            # no end ring resistors to behave like FEMM model
                            study.GetCircuit().CreateWire(X + 2, Y, X + 2, Y - 3)
                            study.GetCircuit().CreateWire(X - 2, Y - 3, X - 2, Y - 6)
                            study.GetCircuit().CreateWire(X + 2, Y - 6, X + 2, Y - 9)
                            study.GetCircuit().CreateWire(X - 2, Y - 9, X - 2, Y - 12)
                            study.GetCircuit().CreateWire(X + 2, Y - 12, X + 2, Y - 15)
                            study.GetCircuit().CreateWire(X - 2, Y, X - 5, Y)
                            study.GetCircuit().CreateWire(X - 5, Y, X - 2, Y - 15)
                            study.GetCircuit().CreateInstance("Ground", X - 5, Y - 2)
                        else:
                            raise Exception(
                                'Not implemented error: pole-specific rotor winding circuit for %d poles are not implemented' % (
                                    im.DriveW_poles))
                else:
                    raise Exception(
                        'Not implemented error: pole-specific rotor winding circuit for %d poles and non-zero End_Ring_Resistance are not implemented' % (
                            im.DriveW_poles))

                for i in range(0, int(self.no_slot_per_pole)):
                    natural_i = i + 1
                    if self.DriveW_poles == 2:
                        study.GetCondition("CdctCon %d" % (natural_i)).SetLink(
                            "Conductor%s1" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + self.no_slot_per_pole)).SetLink(
                            "Conductor%s2" % (rotor_phase_name_list[i]))
                    elif self.DriveW_poles == 4:
                        study.GetCondition("CdctCon %d" % (natural_i)).SetLink(
                            "Conductor%s1" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + self.no_slot_per_pole)).SetLink(
                            "Conductor%s2" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 2 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s3" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 3 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s4" % (rotor_phase_name_list[i]))
                    elif self.DriveW_poles == 6:
                        study.GetCondition("CdctCon %d" % (natural_i)).SetLink(
                            "Conductor%s1" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + self.no_slot_per_pole)).SetLink(
                            "Conductor%s2" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 2 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s3" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 3 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s4" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 4 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s5" % (rotor_phase_name_list[i]))
                        study.GetCondition("CdctCon %d" % (natural_i + 5 * self.no_slot_per_pole)).SetLink(
                            "Conductor%s6" % (rotor_phase_name_list[i]))
                    else:
                        raise Exception('Not implemented for poles %d' % (self.DriveW_poles))
            else:  # Caged rotor circuit
                dyn_circuit = study.GetCircuit().CreateDynamicCircuit("Cage")
                dyn_circuit.SetValue("AntiPeriodic", False)
                dyn_circuit.SetValue("Bars", int(self.Qr))
                dyn_circuit.SetValue("EndringResistance", self.End_Ring_Resistance)
                dyn_circuit.SetValue("GroupCondition", True)
                dyn_circuit.SetValue("GroupName", "CdctCon_Group")
                dyn_circuit.SetValue("UseInductance", False)
                dyn_circuit.Submit("Cage1", 23, 2)
                study.GetCircuit().CreateInstance("Ground", 25, 1)

    def fea_bearingless_induction(self, im_template, x_denorm, counter, counter_loop):
        logger = logging.getLogger(__name__)
        print('Run FEA for individual #%d' % (counter))

        # get local design variant
        im_variant = population.bearingless_induction_motor_design.local_design_variant(im_template, 0, counter,
                                                                                        x_denorm)

        # print('::', im_template.Radius_OuterRotor, im_template.Width_RotorSlotOpen)
        # print('::', im_variant.Radius_OuterRotor, im_variant.Width_RotorSlotOpen)
        # quit()

        if counter_loop == 1:
            im_variant.name = 'ind%d' % (counter)
        else:
            im_variant.name = 'ind%d-redo%d' % (counter, counter_loop)
        # im_variant.spec = im_template.spec
        self.im_variant = im_variant
        self.femm_solver = FEMM_Solver.FEMM_Solver(self.im_variant, flag_read_from_jmag=False, freq=50)  # eddy+static
        im = None

        if counter_loop == 1:
            self.project_name = 'proj%d' % (counter)
        else:
            self.project_name = 'proj%d-redo%d' % (counter, counter_loop)
        self.expected_project_file = self.output_dir + "%s.jproj" % (self.project_name)

        original_study_name = im_variant.name + "Freq"
        tran2tss_study_name = im_variant.name + 'Tran2TSS'

        self.dir_femm_temp = self.output_dir + 'femm_temp/'
        self.femm_output_file_path = self.dir_femm_temp + original_study_name + '.csv'

        # self.jmag_control_state = False

        # local scripts
        def open_jmag(expected_project_file_path):
            if self.app is None:
                # app = win32com.client.Dispatch('designer.Application.181')
                app = win32com.client.Dispatch('designer.Application.181')
                # app = win32com.client.gencache.EnsureDispatch('designer.Application.171') # https://stackoverflow.com/questions/50127959/win32-dispatch-vs-win32-gencache-in-python-what-are-the-pros-and-cons

                if self.fea_config_dict['designer.Show'] == True:
                    app.Show()
                else:
                    app.Hide()
                # app.Quit()
                self.app = app  # means that the JMAG Designer is turned ON now.
                self.bool_run_in_JMAG_Script_Editor = False

                def add_steel(self):
                    print('[First run on this computer detected]', im_template.spec_input_dict['Steel'],
                          'is added to jmag material library.')
                    import population
                    if 'M15' in im_template.spec_input_dict['Steel']:
                        population.add_M1xSteel(self.app, self.fea_config_dict['dir.parent'], steel_name="M-15 Steel")
                    elif 'M19' in im_template.spec_input_dict['Steel']:
                        population.add_M1xSteel(self.app, self.fea_config_dict['dir.parent'])
                    elif 'Arnon5' == im_template.spec_input_dict['Steel']:
                        population.add_Arnon5(self.app, self.fea_config_dict['dir.parent'])

                # too avoid tons of the same material in JAMG's material library
                fname = self.fea_config_dict['dir.parent'] + '.jmag_state.txt'
                if not os.path.exists(fname):
                    with open(fname, 'w') as f:
                        f.write(self.fea_config_dict['pc_name'] + '/' + im_template.spec_input_dict['Steel'] + '\n')
                    add_steel(self)
                else:
                    with open(fname, 'r') as f:
                        for line in f.readlines():
                            if self.fea_config_dict['pc_name'] + '/' + im_template.spec_input_dict['Steel'] not in line:
                                add_steel(self)

            else:
                app = self.app

            print(expected_project_file_path)
            if os.path.exists(expected_project_file_path):
                print(
                    'JMAG project exists already. I learned my lessions. I will NOT delete it but create a new one with a different name instead.')
                # os.remove(expected_project_file_path)
                attempts = 2
                temp_path = expected_project_file_path[:-len('.jproj')] + 'attempts%d.jproj' % (attempts)
                while os.path.exists(temp_path):
                    attempts += 1
                    temp_path = expected_project_file_path[:-len('.jproj')] + 'attempts%d.jproj' % (attempts)

                expected_project_file_path = temp_path

            app.NewProject("Untitled")
            app.SaveAs(expected_project_file_path)
            logger.debug('Create JMAG project file: %s' % (expected_project_file_path))

            return app

        def draw_jmag(app):
            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            # Draw the model in JMAG Designer
            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            print("Inside draw_jmag")
            DRAW_SUCCESS = self.draw_jmag_induction(app,
                                                    counter,
                                                    im_variant,
                                                    im_variant.name)
            if DRAW_SUCCESS == 0:
                # TODO: skip this model and its evaluation
                cost_function = 99999  # penalty
                logging.getLogger(__name__).warn('Draw Failed for %s-%s\nCost function penalty = %g.%s',
                                                 self.project_name, im_variant.name, cost_function,
                                                 self.im_variant.show(toString=True))
                raise Exception(
                    'Draw Failed: Are you working on the PC? Sometime you by mistake operate in the JMAG Geometry Editor, then it fails to draw.')
                return None
            elif DRAW_SUCCESS == -1:
                raise Exception(' DRAW_SUCCESS == -1:')

            # JMAG
            if app.NumModels() >= 1:
                model = app.GetModel(im_variant.name)
            else:
                logger.error('there is no model yet for %s' % (im_variant.name))
                raise Exception('why is there no model yet? %s' % (im_variant.name))
            return model

        def rotating_static_FEA():

            # wait for femm to finish, and get your slip of breakdown
            new_fname = self.dir_femm_temp + original_study_name + '.csv'
            with open(new_fname, 'r') as f:
                data = f.readlines()
                freq = float(data[0][:-1])
                torque = float(data[1][:-1])
            slip_freq_breakdown_torque, breakdown_torque, breakdown_force = freq, torque, None
            # Now we have the slip, set it up!
            im_variant.update_mechanical_parameters(slip_freq_breakdown_torque)  # do this for records only

            # Must run this script after slip_freq_breakdown_torque is known to get new results, but if results are present, it is okay to use this script for post-processing.
            if not self.femm_solver.has_results():
                print('run_rotating_static_FEA')
                # utility.blockPrint()
                self.femm_solver.run_rotating_static_FEA()
                self.femm_solver.parallel_solve()
                # utility.enablePrint()

            # collecting parasolve with post-process
            # wait for .ans files
            # data_femm_solver = self.femm_solver.show_results_static(bool_plot=False) # this will wait as well?
            while not self.femm_solver.has_results():
                print(clock_time())
                sleep(3)
            results_dict = {}
            for f in [f for f in os.listdir(self.femm_solver.dir_run) if 'static_results' in f]:
                data = np.loadtxt(self.femm_solver.dir_run + f, unpack=True, usecols=(0, 1, 2, 3))
                for i in range(len(data[0])):
                    results_dict[data[0][i]] = (data[1][i], data[2][i], data[3][i])
            keys_without_duplicates = [key for key, item in results_dict.items()]
            keys_without_duplicates.sort()
            with open(self.femm_solver.dir_run + "no_duplicates.txt", 'w') as fw:
                for key in keys_without_duplicates:
                    fw.writelines(
                        '%g %g %g %g\n' % (key, results_dict[key][0], results_dict[key][1], results_dict[key][2]))
            data_femm_solver = np.array([keys_without_duplicates,
                                         [results_dict[key][0] for key in keys_without_duplicates],
                                         [results_dict[key][1] for key in keys_without_duplicates],
                                         [results_dict[key][2] for key in keys_without_duplicates]])
            # print(data_femm_solver)
            # from pylab import plt
            # plt.figure(); plt.plot(data_femm_solver[0])
            # plt.figure(); plt.plot(data_femm_solver[1])
            # plt.figure(); plt.plot(data_femm_solver[2])
            # plt.figure(); plt.plot(data_femm_solver[3])
            # plt.show()
            # quit()
            return data_femm_solver

        def rotating_eddy_current_FEA(im, app, model):

            # Freq Sweeping for break-down Torque Slip
            # remember to export the B data using subroutine
            # and check export table results only
            study = im.add_study(app, model, self.dir_csv_output_folder, choose_study_type='frequency')

            # Freq Study: you can choose to not use JMAG to find the breakdown slip.
            # Option 1: you can set im.slip_freq_breakdown_torque by FEMM Solver
            # Option 2: Use JMAG to sweeping the frequency
            # Does study has results?
            if study.AnyCaseHasResult():
                slip_freq_breakdown_torque, breakdown_torque, breakdown_force = self.check_csv_results(study.GetName())
            else:
                # mesh
                im.add_mesh(study, model)

                # Export Image
                # for i in range(app.NumModels()):
                #     app.SetCurrentModel(i)
                #     model = app.GetCurrentModel()
                #     app.ExportImage(r'D:\Users\horyc\OneDrive - UW-Madison\pop\run#10/' + model.GetName() + '.png')
                app.View().ShowAllAirRegions()
                # app.View().ShowMeshGeometry() # 2nd btn
                app.View().ShowMesh()  # 3rn btn
                app.View().Zoom(3)
                app.View().Pan(-im.Radius_OuterRotor, 0)
                app.ExportImageWithSize('./' + model.GetName() + '.png', 2000, 2000)
                app.View().ShowModel()  # 1st btn. close mesh view, and note that mesh data will be deleted because only ouput table results are selected.

                # run
                study.RunAllCases()
                app.Save()

                def check_csv_results(dir_csv_output_folder, study_name, returnBoolean=False,
                                      file_suffix='_torque.csv'):  # '_iron_loss_loss.csv'
                    # print self.dir_csv_output_folder + study_name + '_torque.csv'
                    if not os.path.exists(dir_csv_output_folder + study_name + file_suffix):
                        if returnBoolean == False:
                            print('Nothing is found when looking into:',
                                  dir_csv_output_folder + study_name + file_suffix)
                            return None
                        else:
                            return False
                    else:
                        if returnBoolean == True:
                            return True

                    try:
                        # check csv results
                        l_slip_freq = []
                        l_TorCon = []
                        l_ForCon_X = []
                        l_ForCon_Y = []

                        with open(dir_csv_output_folder + study_name + '_torque.csv', 'r') as f:
                            for ind, row in enumerate(utility.csv_row_reader(f)):
                                if ind >= 5:
                                    try:
                                        float(row[0])
                                    except:
                                        continue
                                    l_slip_freq.append(float(row[0]))
                                    l_TorCon.append(float(row[1]))

                        with open(dir_csv_output_folder + study_name + '_force.csv', 'r') as f:
                            for ind, row in enumerate(utility.csv_row_reader(f)):
                                if ind >= 5:
                                    try:
                                        float(row[0])
                                    except:
                                        continue
                                    # l_slip_freq.append(float(row[0]))
                                    l_ForCon_X.append(float(row[1]))
                                    l_ForCon_Y.append(float(row[2]))

                        breakdown_force = max(np.sqrt(np.array(l_ForCon_X) ** 2 + np.array(l_ForCon_Y) ** 2))

                        index, breakdown_torque = utility.get_index_and_max(l_TorCon)
                        slip_freq_breakdown_torque = l_slip_freq[index]
                        return slip_freq_breakdown_torque, breakdown_torque, breakdown_force
                    except NameError as e:
                        logger = logging.getLogger(__name__)
                        logger.error('No CSV File Found.', exc_info=True)
                        raise e

                # evaluation based on the csv results
                print(':::ZZZ', self.dir_csv_output_folder)
                slip_freq_breakdown_torque, breakdown_torque, breakdown_force = check_csv_results(
                    self.dir_csv_output_folder, study.GetName())

            # this will be used for other duplicated studies
            original_study_name = study.GetName()
            im.csv_previous_solve = self.dir_csv_output_folder + original_study_name + '_circuit_current.csv'
            im.update_mechanical_parameters(slip_freq_breakdown_torque, syn_freq=im.DriveW_Freq)

            # EC Rotate: Rotate the rotor to find the ripples in force and torque # 不关掉这些云图，跑第二个study的时候，JMAG就挂了：app.View().SetVectorView(False); app.View().SetFluxLineView(False); app.View().SetContourView(False)
            ecrot_study_name = original_study_name + "-FFVRC"
            casearray = [0 for i in range(1)]
            casearray[0] = 1
            model.DuplicateStudyWithCases(original_study_name, ecrot_study_name, casearray)

            app.SetCurrentStudy(ecrot_study_name)
            study = app.GetCurrentStudy()

            divisions_per_slot_pitch = 24  # self.fea_config_dict['ec_rotate_divisions_per_slot_pitch']  # 24
            study.GetStep().SetValue("Step", divisions_per_slot_pitch)
            study.GetStep().SetValue("StepType", 0)
            study.GetStep().SetValue("FrequencyStep", 0)
            study.GetStep().SetValue("Initialfrequency", slip_freq_breakdown_torque)

            # study.GetCondition(u"RotCon").SetValue(u"MotionGroupType", 1)
            study.GetCondition("RotCon").SetValue("Displacement", + 360.0 / im.Qr / divisions_per_slot_pitch)

            # https://www2.jmag-international.com/support/en/pdf/JMAG-Designer_Ver.17.1_ENv3.pdf
            study.GetStudyProperties().SetValue("DirectSolverType", 1)

            # model.RestoreCadLink()
            study.Run()
            app.Save()
            # model.CloseCadLink()

        def transient_FEA_as_reference(im, slip_freq_breakdown_torque):
            # Transient Reference
            tranRef_study_name = "TranRef"
            if model.NumStudies() < 4:
                model.DuplicateStudyWithType(tran2tss_study_name, "Transient2D", tranRef_study_name)
                app.SetCurrentStudy(tranRef_study_name)
                study = app.GetCurrentStudy()

                # 将一个滑差周期和十个同步周期，分成 400 * end_point / (1.0/im.DriveW_Freq) 份。
                end_point = 0.5 / slip_freq_breakdown_torque + 10.0 / im.DriveW_Freq
                # Pavel Ponomarev 推荐每个电周期400~600个点来捕捉槽效应。
                division = self.fea_config_dict['designer.TranRef-StepPerCycle'] * end_point / (
                            1.0 / im.DriveW_Freq)  # int(end_point * 1e4)
                # end_point = division * 1e-4
                study.GetStep().SetValue("Step", division + 1)
                study.GetStep().SetValue("StepType", 1)  # regular inverval
                study.GetStep().SetValue("StepDivision", division)
                study.GetStep().SetValue("EndPoint", end_point)

                # https://www2.jmag-international.com/support/en/pdf/JMAG-Designer_Ver.17.1_ENv3.pdf
                study.GetStudyProperties().SetValue("DirectSolverType", 1)

                study.RunAllCases()
                app.Save()

        # debug for tia-iemdc-ecce-2019
        # data_femm_solver = rotating_static_FEA()
        # from show_results_iemdc19 import show_results_iemdc19
        # show_results_iemdc19(   self.dir_csv_output_folder,
        #                         im_variant,
        #                         femm_solver_data=data_femm_solver,
        #                         femm_rotor_current_function=self.femm_solver.get_rotor_current_function()
        #                     )
        # quit()

        if self.fea_config_dict['bool_re_evaluate'] == False:

            # this should be summoned even before initializing femm, and it will decide whether the femm results are reliable
            app = open_jmag(self.expected_project_file)  # will set self.jmag_control_state to True

            ################################################################
            # Begin from where left: Frequency Study
            ################################################################
            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            # Eddy Current Solver for Breakdown Torque and Slip
            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            # check for existing results
            if os.path.exists(self.femm_output_file_path):
                # for file in os.listdir(self.dir_femm_temp):
                #     if original_study_name in file:
                #         print('----------', original_study_name, file)
                print('Remove legacy femm output files @ %s' % (self.femm_output_file_path))
                os.remove(self.femm_output_file_path)
                os.remove(self.femm_output_file_path[:-4] + '.fem')
                # quit()
                # quit()

            # delete me
            # model = draw_jmag(app)

            # At this point, no results exist from femm.
            print('Run greedy_search_for_breakdown_slip...')
            femm_tic = clock_time()
            # self.femm_solver.__init__(im_variant, flag_read_from_jmag=False, freq=50.0)
            if im_variant.DriveW_poles == 4 and self.fea_config_dict['femm.use_fraction'] == True:
                print('FEMM model only solves for a fraction of 2.\n' * 3)
                self.femm_solver.greedy_search_for_breakdown_slip(self.dir_femm_temp, original_study_name,
                                                                  bool_run_in_JMAG_Script_Editor=self.bool_run_in_JMAG_Script_Editor,
                                                                  fraction=2)
            else:
                # p >= 3 is not tested so do not use fraction for now
                self.femm_solver.greedy_search_for_breakdown_slip(self.dir_femm_temp, original_study_name,
                                                                  bool_run_in_JMAG_Script_Editor=self.bool_run_in_JMAG_Script_Editor,
                                                                  fraction=1)  # 转子导条必须形成通路

            ################################################################
            # Begin from where left: Transient Study
            ################################################################
            model = draw_jmag(app)

            # EC-Rotate
            # rotating_eddy_current_FEA(im_variant, app, model)

            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            # TranFEAwi2TSS for ripples and iron loss
            # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
            # add or duplicate study for transient FEA denpending on jmag_run_list
            # FEMM+JMAG (注意，这里我们用50Hz作为滑差频率先设置起来，等拿到breakdown slip freq的时候，再更新变量slip和study properties的时间。)
            study = im_variant.add_TranFEAwi2TSS_study(50.0, app, model, self.dir_csv_output_folder,
                                                       tran2tss_study_name, logger)
            app.SetCurrentStudy(tran2tss_study_name)
            study = app.GetCurrentStudy()
            self.mesh_study(im_variant, app, model, study)

            # wait for femm to finish, and get your slip of breakdown
            slip_freq_breakdown_torque, breakdown_torque, breakdown_force = self.femm_solver.wait_greedy_search(
                femm_tic)

            # Now we have the slip, set it up!
            im_variant.update_mechanical_parameters(slip_freq_breakdown_torque)  # do this for records only
            if im_variant.the_slip != slip_freq_breakdown_torque / im_variant.DriveW_Freq:
                raise Exception('Check update_mechanical_parameters().')
            study.GetDesignTable().GetEquation("slip").SetExpression("%g" % (im_variant.the_slip))
            if True:
                number_of_steps_2ndTSS = self.fea_config_dict['designer.number_of_steps_2ndTSS']
                DM = app.GetDataManager()
                DM.CreatePointArray("point_array/timevsdivision", "SectionStepTable")
                refarray = [[0 for i in range(3)] for j in range(3)]
                refarray[0][0] = 0
                refarray[0][1] = 1
                refarray[0][2] = 50
                refarray[1][0] = 0.5 / slip_freq_breakdown_torque  # 0.5 for 17.1.03l # 1 for 17.1.02y
                refarray[1][1] = number_of_steps_2ndTSS  # 16 for 17.1.03l #32 for 17.1.02y
                refarray[1][2] = 50
                refarray[2][0] = refarray[1][0] + 0.5 / im_variant.DriveW_Freq  # 0.5 for 17.1.03l
                refarray[2][1] = number_of_steps_2ndTSS  # also modify range_ss! # don't forget to modify below!
                refarray[2][2] = 50
                DM.GetDataSet("SectionStepTable").SetTable(refarray)
                number_of_total_steps = 1 + 2 * number_of_steps_2ndTSS  # [Double Check] don't forget to modify here!
                study.GetStep().SetValue("Step", number_of_total_steps)
                study.GetStep().SetValue("StepType", 3)
                study.GetStep().SetTableProperty("Division", DM.GetDataSet("SectionStepTable"))

            # static FEA solver with FEMM (need eddy current FEA results)
            # print('::', im_variant.Omega, im_variant.Omega/2/np.pi)
            # print('::', femm_solver.im.Omega, femm_solver.im.Omega/2/np.pi)
            # quit()
            # rotating_static_FEA()

            # debug JMAG circuit
            # app.Save()
            # quit()

            # Run JMAG study
            self.run_study(im_variant, app, study, clock_time())

            # export Voltage if field data exists.
            if self.fea_config_dict['delete_results_after_calculation'] == False:
                # Export Circuit Voltage
                ref1 = app.GetDataManager().GetDataSet("Circuit Voltage")
                app.GetDataManager().CreateGraphModel(ref1)
                app.GetDataManager().GetGraphModel("Circuit Voltage").WriteTable(
                    self.dir_csv_output_folder + im_variant.name + "_EXPORT_CIRCUIT_VOLTAGE.csv")

            # TranRef
            # transient_FEA_as_reference(im_variant, slip_freq_breakdown_torque)

        else:
            # FEMM的转子电流，哎，是个麻烦事儿。。。

            self.femm_solver.vals_results_rotor_current = []

            new_fname = self.dir_femm_temp + original_study_name + '.csv'
            try:
                with open(new_fname, 'r') as f:
                    buf = f.readlines()
                    for idx, el in enumerate(buf):
                        if idx == 0:
                            slip_freq_breakdown_torque = float(el)
                            continue
                        if idx == 1:
                            breakdown_torque = float(el)
                            continue
                        if idx == 2:
                            self.femm_solver.stator_slot_area = float(el)
                            continue
                        if idx == 3:
                            self.femm_solver.rotor_slot_area = float(el)
                            continue
                        # print(el)
                        # print(el.split(','))
                        temp = el.split(',')
                        self.femm_solver.vals_results_rotor_current.append(float(temp[0]) + 1j * float(temp[1]))
                        # print(self.femm_solver.vals_results_rotor_current)
                self.dirty_backup_stator_slot_area = self.femm_solver.stator_slot_area
                self.dirty_backup_rotor_slot_area = self.femm_solver.rotor_slot_area
                self.dirty_backup_vals_results_rotor_current = self.femm_solver.vals_results_rotor_current
            except FileNotFoundError as error:
                print(error)
                print('Use dirty_backup to continue...')  # 有些时候，不知道为什么femm的结果文件（.csv）没了，这时候曲线救国，凑活一下吧
                self.femm_solver.stator_slot_area = self.dirty_backup_stator_slot_area
                self.femm_solver.rotor_slot_area = self.dirty_backup_rotor_slot_area
                self.femm_solver.vals_results_rotor_current = self.dirty_backup_vals_results_rotor_current

                # 电机的电流值取决于槽的面积。。。。
            THE_mm2_slot_area = self.femm_solver.stator_slot_area * 1e6

            if 'VariableStatorSlotDepth' in self.fea_config_dict['which_filter']:
                # set DriveW_CurrentAmp using the calculated stator slot area.
                print('[A]: DriveW_CurrentAmp is updated.')

                # 槽深变化，电密不变，所以电流也会变化。
                CurrentAmp_in_the_slot = THE_mm2_slot_area * im_variant.fill_factor * im_variant.Js * 1e-6 * np.sqrt(2)
                CurrentAmp_per_conductor = CurrentAmp_in_the_slot / im_variant.DriveW_zQ
                CurrentAmp_per_phase = CurrentAmp_per_conductor * im_variant.wily.number_parallel_branch  # 跟几层绕组根本没关系！除以zQ的时候，就已经变成每根导体的电流了。
                variant_DriveW_CurrentAmp = CurrentAmp_per_phase  # this current amp value is for non-bearingless motor

                im_variant.CurrentAmp_per_phase = CurrentAmp_per_phase

                im_variant.DriveW_CurrentAmp = self.fea_config_dict['TORQUE_CURRENT_RATIO'] * variant_DriveW_CurrentAmp
                im_variant.BeariW_CurrentAmp = self.fea_config_dict[
                                                   'SUSPENSION_CURRENT_RATIO'] * variant_DriveW_CurrentAmp

                slot_area_utilizing_ratio = (
                                                        im_variant.DriveW_CurrentAmp + im_variant.BeariW_CurrentAmp) / im_variant.CurrentAmp_per_phase
                print('---Heads up! slot_area_utilizing_ratio is', slot_area_utilizing_ratio)

                print('---Variant CurrentAmp_in_the_slot =', CurrentAmp_in_the_slot)
                print('---variant_DriveW_CurrentAmp = CurrentAmp_per_phase =', variant_DriveW_CurrentAmp)
                print('---im_variant.DriveW_CurrentAmp =', im_variant.DriveW_CurrentAmp)
                print('---im_variant.BeariW_CurrentAmp =', im_variant.BeariW_CurrentAmp)
                print('---TORQUE_CURRENT_RATIO:', self.fea_config_dict['TORQUE_CURRENT_RATIO'])
                print('---SUSPENSION_CURRENT_RATIO:', self.fea_config_dict['SUSPENSION_CURRENT_RATIO'])

        ################################################################
        # Load data for cost function evaluation
        ################################################################
        im_variant.results_to_be_unpacked = results_to_be_unpacked = utility.build_str_results(self.axeses, im_variant,
                                                                                               self.project_name,
                                                                                               tran2tss_study_name,
                                                                                               self.dir_csv_output_folder,
                                                                                               self.fea_config_dict,
                                                                                               self.femm_solver,
                                                                                               machine_type='IM')
        if results_to_be_unpacked is not None:
            if self.fig_main is not None:
                try:
                    self.fig_main.savefig(self.output_dir + im_variant.name + 'results.png', dpi=150)
                except Exception as e:
                    print('Directory exists?', self.output_dir + im_variant.name + 'results.png')
                    print(e)
                    print('\n\n\nIgnore error and continue.')
                finally:
                    utility.pyplot_clear(self.axeses)
            # show()
            return im_variant
        else:
            raise Exception('results_to_be_unpacked is None.')
        # winding analysis? 之前的python代码利用起来啊
        # 希望的效果是：设定好一个设计，马上进行运行求解，把我要看的数据都以latex报告的形式呈现出来。
        # OP_PS_Qr36_M19Gauge29_DPNV_NoEndRing.jproj

    # def draw_jmag(app):
    #     # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    #     # Draw the model in JMAG Designer
    #     # ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    #     print("Inside draw_jmag")
    #     DRAW_SUCCESS = self.draw_jmag_induction(app,
    #                                             counter,
    #                                             im_variant,
    #                                             im_variant.name)
    #     if DRAW_SUCCESS == 0:
    #         # TODO: skip this model and its evaluation
    #         cost_function = 99999  # penalty
    #         logging.getLogger(__name__).warn('Draw Failed for %s-%s\nCost function penalty = %g.%s', self.project_name,
    #                                          im_variant.name, cost_function, self.im_variant.show(toString=True))
    #         raise Exception(
    #             'Draw Failed: Are you working on the PC? Sometime you by mistake operate in the JMAG Geometry Editor, then it fails to draw.')
    #         return None
    #     elif DRAW_SUCCESS == -1:
    #         raise Exception(' DRAW_SUCCESS == -1:')
    #
    #     # JMAG
    #     if app.NumModels() >= 1:
    #         model = app.GetModel(im_variant.name)
    #     else:
    #         logger.error('there is no model yet for %s' % (im_variant.name))
    #         raise Exception('why is there no model yet? %s' % (im_variant.name))
    #     return model

    def draw_jmag_induction(self, app, individual_index, im_variant, model_name, bool_trimDrawer_or_vanGogh=True,
                            doNotRotateCopy=False):
        print('Inside draw jmag induction')
        if individual_index == -1:  # 后处理是-1
            print('Draw model for post-processing')
            if individual_index + 1 + 1 <= app.NumModels():
                logger = logging.getLogger(__name__)
                logger.debug('The model already exists for individual with index=%d. Skip it.', individual_index)
                return -1  # the model is already drawn

        elif individual_index + 1 <= app.NumModels():  # 一般是从零起步
            logger = logging.getLogger(__name__)
            logger.debug('The model already exists for individual with index=%d. Skip it.', individual_index)
            return -1  # the model is already drawn

        # open JMAG Geometry Editor
        app.LaunchGeometryEditor()
        geomApp = app.CreateGeometryEditor()
        # geomApp.Show()
        geomApp.NewDocument()
        doc = geomApp.GetDocument()
        ass = doc.GetAssembly()

        # draw parts
        print('Before try statement')
        try:
            print('Bool Trimmer', bool_trimDrawer_or_vanGogh)
            if bool_trimDrawer_or_vanGogh:
                print('Before population import')
                # from . import population
                print('Before TrimDrawer')
                d = population.TrimDrawer(self.machine_variant)  # 传递的是地址哦
                print('After TrimDrawer')
                d.doc, d.ass = doc, ass
                d.plot_shaft("Shaft")

                d.plot_rotorCore("Rotor Core")
                d.plot_cage("Cage")

                d.plot_statorCore("Stator Core")

                d.plot_coil("Coil")
                # d.plot_airWithinRotorSlots(u"Air Within Rotor Slots")

                if 'VariableStatorSlotDepth' in self.configuration['which_filter']:
                    # set DriveW_CurrentAmp using the calculated stator slot area.
                    print('[A]: DriveW_CurrentAmp is updated.')

                    # 槽深变化，电密不变，所以电流也会变化。
                    CurrentAmp_in_the_slot = d.mm2_slot_area * im_variant.fill_factor * im_variant.Js * 1e-6 * np.sqrt(
                        2)
                    CurrentAmp_per_conductor = CurrentAmp_in_the_slot / im_variant.DriveW_zQ

                    CurrentAmp_per_phase = CurrentAmp_per_conductor * im_variant.number_parallel_branch  # 跟几层绕组根本没关系！除以zQ的时候，就已经变成每根导体的电流了。
                    variant_DriveW_CurrentAmp = CurrentAmp_per_phase  # this current amp value is for non-bearingless motor

                    im_variant._machine_parameter_dict["CurrentAmp_per_phase"] = CurrentAmp_per_phase

                    im_variant._machine_parameter_dict["DriveW_CurrentAmp"] = im_variant.fea_config_dict[
                                                                                  'TORQUE_CURRENT_RATIO'] * variant_DriveW_CurrentAmp
                    im_variant._machine_parameter_dict["BeariW_CurrentAmp"] = im_variant.fea_config_dict[
                                                                                  'SUSPENSION_CURRENT_RATIO'] * variant_DriveW_CurrentAmp

                    from pprint import pprint
                    pprint(vars(im_variant))

                    slot_area_utilizing_ratio = (
                                                        im_variant.DriveW_CurrentAmp + im_variant.BeariW_CurrentAmp) / im_variant.CurrentAmp_per_phase
                    print('---Heads up! slot_area_utilizing_ratio is', slot_area_utilizing_ratio)

                    print('---Variant CurrentAmp_in_the_slot =', CurrentAmp_in_the_slot)
                    print('---variant_DriveW_CurrentAmp = CurrentAmp_per_phase =', variant_DriveW_CurrentAmp)
                    print('---im_variant.DriveW_CurrentAmp =', im_variant.DriveW_CurrentAmp)
                    print('---im_variant.BeariW_CurrentAmp =', im_variant.BeariW_CurrentAmp)
                    print('---TORQUE_CURRENT_RATIO:', im_variant.fea_config_dict['TORQUE_CURRENT_RATIO'])
                    print('---SUSPENSION_CURRENT_RATIO:', im_variant.fea_config_dict['SUSPENSION_CURRENT_RATIO'])

            else:
                print("Call Vangogh at em_analyzer")
                d = VanGogh_JMAG(self.machine_variant, doNotRotateCopy=doNotRotateCopy)
                print("Call Vangogh at em_analyzer1")
                d.doc, d.ass = doc, ass
                d.draw_model()
            self.d = d
        except Exception as e:
            print('See log file to plotting error.')
            logger = logging.getLogger(__name__)
            logger.error('The drawing is terminated. Please check whether the specified bounds are proper.',
                         exc_info=True)

            raise e

            # print 'Draw Failed'
            # if self.pc_name == 'Y730':
            #     # and send the email to hory chen
            #     raise e

            # or you can skip this model and continue the optimization!
            return False  # indicating the model cannot be drawn with the script.

        # Import Model into Designer
        doc.SaveModel(True)  # True=on : Project is also saved.
        model = app.GetCurrentModel()  # model = app.GetModel(u"IM_DEMO_1")
        model.SetName(model_name)
        model.SetDescription(im_variant.fea_config_dict['model_name_prefix'])

        # if doNotRotateCopy:
        #     im_variant.pre_process_structural(app, d.listKeyPoints)
        # else:
        #     im_variant.pre_process(app)
        #
        # model.CloseCadLink()  # this is essential if you want to create a series of models
        # return True
