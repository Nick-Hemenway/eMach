from time import time as clock_time
import os
from .FEMM_Solver import FEMM_Solver
import win32com.client
import logging
logger = logging.getLogger(__name__)
# import population



class IM_EM_Analysis():

    def __init__(self, configuration):
        self.configuration = configuration
        self.machine_variant = None
        self.operating_point = None

    def analyze(self, problem, counter = 0):
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
                from . import population
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
                    CurrentAmp_per_phase = CurrentAmp_per_conductor * im_variant.wily.number_parallel_branch  # 跟几层绕组根本没关系！除以zQ的时候，就已经变成每根导体的电流了。
                    variant_DriveW_CurrentAmp = CurrentAmp_per_phase  # this current amp value is for non-bearingless motor

                    im_variant.CurrentAmp_per_phase = CurrentAmp_per_phase

                    im_variant.DriveW_CurrentAmp = self.fea_config_dict[
                                                       'TORQUE_CURRENT_RATIO'] * variant_DriveW_CurrentAmp
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
        model.SetDescription(im_variant.model_name_prefix + '\n' + im_variant.show(toString=True))

        if doNotRotateCopy:
            im_variant.pre_process_structural(app, d.listKeyPoints)
        else:
            im_variant.pre_process(app)

        model.CloseCadLink()  # this is essential if you want to create a series of models
        return True








            


    


