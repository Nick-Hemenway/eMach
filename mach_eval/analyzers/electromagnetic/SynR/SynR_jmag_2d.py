# Created: 4/13/2023
# Author: Dante Newman

import os
import numpy as np
import pandas as pd
import sys
from time import time as clock_time

# add the directory three levels above this file's directory to path for module import
sys.path.append(os.path.dirname(__file__))

from eMach.mach_cad import model_obj as mo
from eMach.mach_opt import InvalidDesign
from eMach.mach_cad.tools import jmag as JMAG
from eMach.mach_eval.analyzers.electromagnetic.stator_wdg_res import(
    StatorWindingResistanceProblem, StatorWindingResistanceAnalyzer
)

class SynR_EM_Problem:
    def __init__(self, machine, operating_point):
        self.machine = machine
        self.operating_point = operating_point
        self._validate_attr()

    def _validate_attr(self):
        if 'SynR_Machine' in str(type(self.machine)):
            pass
        else:
            raise TypeError("Invalid machine type")

        if 'SynR_Machine_Oper_Pt' in str(type(self.operating_point)):
            pass
        else:
            raise TypeError("Invalid settings type")


class SynR_EM_Analyzer:
    def __init__(self, configuration):
        self.config = configuration

    def analyze(self, problem):
        self.machine_variant = problem.machine
        self.operating_point = problem.operating_point
        ####################################################
        # 01 Setting project name and output folder
        ####################################################
        self.project_name = self.machine_variant.name
        # expected_project_file = self.config.run_folder + "%s_attempts_2.jproj" % self.project_name
        expected_project_file = self.config.run_folder + "%s.jproj" % self.project_name
        
        # Create output folder
        if not os.path.isdir(self.config.jmag_csv_folder):
            os.makedirs(self.config.jmag_csv_folder)

        attempts = 1
        if os.path.exists(expected_project_file):
            print(
                "JMAG project exists already, I will not delete it but create a new one with a different name instead."
            )
            # os.remove(expected_project_file_path)
            attempts = 2
            temp_path = expected_project_file[
                : -len(".jproj")
            ] + "_attempts_%d.jproj" % (attempts)
            while os.path.exists(temp_path):
                attempts += 1
                temp_path = expected_project_file[
                    : -len(".jproj")
                ] + "_attempts_%d.jproj" % (attempts)

            expected_project_file = temp_path

        if attempts > 1:
            self.project_name = self.project_name + "_attempts_%d" % (attempts)

        toolJmag = JMAG.JmagDesigner()

        toolJmag.visible = self.config.jmag_visible
        toolJmag.open(comp_filepath=expected_project_file, length_unit="DimMillimeter", study_type="Transient2D")
        toolJmag.save()

        self.study_name = self.project_name + "_SynR_JMAG"
        self.design_results_folder = (
            self.config.run_folder + "%s_results/" % self.project_name
        )
        if not os.path.isdir(self.design_results_folder):
            os.makedirs(self.design_results_folder)

        ################################################################
        # 02 Run Electromagnetic analysis
        ################################################################
        
        # Draw cross_section
        draw_success = self.draw_machine(toolJmag)
    
        if not draw_success:
            raise InvalidDesign

        toolJmag.doc.SaveModel(False)
        app = toolJmag.jd
        model = app.GetCurrentModel()

        # Pre-processing
        model.SetName(self.project_name)
        model.SetDescription(self.show(self.project_name, toString=True))
        
        valid_design = self.pre_process(model)

        if not valid_design:
            raise InvalidDesign

        # Create transient study with two time step sections
        # model = toolJmag.create_model(self.study_name)
        study = self.add_transient_magnetic_study(app, model, self.config.jmag_csv_folder, self.study_name)
        self.create_custom_material(
            app, self.machine_variant.stator_iron_mat["core_material"]
        )
        self.create_custom_material(
            app, self.machine_variant.coil_mat["coil_material"]
        )
        app.SetCurrentStudy(self.study_name)

        # Mesh study
        self.mesh_study(app, model, study)

        self.run_study(app, study, clock_time())
        # export Voltage if field data exists.
        if not self.config.del_results_after_calc:
            # Export Circuit Voltage
            ref1 = app.GetDataManager().GetDataSet("Circuit Voltage")
            app.GetDataManager().CreateGraphModel(ref1)
            app.GetDataManager().GetGraphModel("Circuit Voltage").WriteTable(
                self.config.jmag_csv_folder
                + self.study_name
                + "_EXPORT_CIRCUIT_VOLTAGE.csv"
            )
        toolJmag.close()
        ####################################################
        # 03 Load FEA output
        ####################################################

        fea_rated_output = self.extract_JMAG_results(
            self.config.jmag_csv_folder, self.study_name
        )

        return fea_rated_output

    @property
    def drive_freq(self):
        drive_freq = self.operating_point.speed / 60 * self.machine_variant.p
        return drive_freq

    @property
    def speed(self):
        return self.operating_point.speed

    @property
    def elec_omega(self):
        return 2 * np.pi * self.drive_freq

    @property
    def z_C(self):
        if len(self.machine_variant.layer_phases) == 1:
            z_C = self.machine_variant.Q / (2 * self.machine_variant.no_of_phases)
        elif len(self.machine_variant.layer_phases) == 2:
            z_C = self.machine_variant.Q / (self.machine_variant.no_of_phases)

        return z_C

    @property
    def stator_resistance(self):
        res_prob = StatorWindingResistanceProblem(
            r_si=self.machine_variant.r_si/1000,
            d_sp=self.machine_variant.d_sp/1000,
            d_st=self.machine_variant.d_st/1000,
            w_st=self.machine_variant.w_st/1000,
            l_st=self.machine_variant.l_st/1000,
            Q=self.machine_variant.Q,
            y=self.machine_variant.pitch,
            z_Q=self.machine_variant.Z_q,
            z_C=self.z_C,
            Kcu=self.machine_variant.Kcu,
            Kov=self.machine_variant.Kov,
            sigma_cond=self.machine_variant.coil_mat["copper_elec_conductivity"],
            slot_area=self.machine_variant.s_slot*1e-6,
        )
        res_analyzer = StatorWindingResistanceAnalyzer()
        stator_resistance = res_analyzer.analyze(res_prob)
        return stator_resistance

    @property
    def R_wdg(self):
        return self.stator_resistance[0]

    @property
    def R_wdg_coil_ends(self):
        return self.stator_resistance[1]

    @property
    def R_wdg_coil_sides(self):
        return self.stator_resistance[2]


    def draw_machine(self, tool):
        ####################################################
        # Adding parts objects
        ####################################################
        self.stator_core = mo.CrossSectInnerRotorStatorPartial(
            name="StatorCore",
            dim_alpha_st=mo.DimDegree(self.machine_variant.alpha_st),
            dim_alpha_so=mo.DimDegree(self.machine_variant.alpha_so),
            dim_r_si=mo.DimMillimeter(self.machine_variant.r_si),
            dim_d_so=mo.DimMillimeter(self.machine_variant.d_so),
            dim_d_sp=mo.DimMillimeter(self.machine_variant.d_sp),
            dim_d_st=mo.DimMillimeter(self.machine_variant.d_st),
            dim_d_sy=mo.DimMillimeter(self.machine_variant.d_sy),
            dim_w_st=mo.DimMillimeter(self.machine_variant.w_st),
            dim_r_st=mo.DimMillimeter(0),
            dim_r_sf=mo.DimMillimeter(0),
            dim_r_sb=mo.DimMillimeter(0),
            Q=self.machine_variant.Q,
            location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)],
            theta=mo.DimRadian(0)),
            )

        self.winding_layer1 = mo.CrossSectInnerRotorStatorRightSlot(
            name="WindingLayer1",
            stator_core=self.stator_core,
            location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)],
            theta=mo.DimRadian(0)),
            )

        self.winding_layer2 = mo.CrossSectInnerRotorStatorLeftSlot(
            name="WindingLayer2",
            stator_core=self.stator_core,
            location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)],
            theta=mo.DimRadian(0)),
            )

        self.rotor_core = mo.CrossSectFluxBarrierRotor(
            name="RotorCore",
            dim_alpha_b=mo.DimDegree(self.machine_variant.alpha_b),
            dim_r_ri=mo.DimMillimeter(self.machine_variant.r_ri),
            dim_r_ro=mo.DimMillimeter(self.machine_variant.r_ro),
            dim_r_f1=mo.DimMillimeter(self.machine_variant.r_f1),
            dim_r_f2=mo.DimMillimeter(self.machine_variant.r_f2),
            dim_r_f3=mo.DimMillimeter(self.machine_variant.r_f3),
            dim_d_r1=mo.DimMillimeter(self.machine_variant.d_r1),
            dim_d_r2=mo.DimMillimeter(self.machine_variant.d_r2),
            dim_d_r3=mo.DimMillimeter(self.machine_variant.d_r3),
            dim_w_b1=mo.DimMillimeter(self.machine_variant.w_b1),
            dim_w_b2=mo.DimMillimeter(self.machine_variant.w_b2),
            dim_w_b3=mo.DimMillimeter(self.machine_variant.w_b3),
            dim_l_b1=mo.DimMillimeter(self.machine_variant.l_b1),
            dim_l_b2=mo.DimMillimeter(self.machine_variant.l_b2),
            dim_l_b3=mo.DimMillimeter(self.machine_variant.l_b3),
            dim_l_b4=mo.DimMillimeter(self.machine_variant.l_b4),
            dim_l_b5=mo.DimMillimeter(self.machine_variant.l_b5),
            dim_l_b6=mo.DimMillimeter(self.machine_variant.l_b6),
            p=2,
            location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)]),
            )

        self.shaft = mo.CrossSectHollowCylinder(
            name="Shaft",
            dim_t=mo.DimMillimeter(self.machine_variant.r_ri),
            dim_r_o=mo.DimMillimeter(self.machine_variant.r_ri),
            location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)]),
            )


        self.comp_stator_core = mo.Component(
            name="StatorCore",
            cross_sections=[self.stator_core],
            material=mo.MaterialGeneric(name=self.machine_variant.stator_iron_mat["core_material"], color=r"#808080"),
            make_solid=mo.MakeExtrude(location=mo.Location3D(), 
                    dim_depth=mo.DimMillimeter(self.machine_variant.l_st)),
            )

        self.comp_winding_layer1 = mo.Component(
            name="WindingLayer1",
            cross_sections=[self.winding_layer1],
            material=mo.MaterialGeneric(name=self.machine_variant.coil_mat["coil_material"]),
            make_solid=mo.MakeExtrude(location=mo.Location3D(), 
            dim_depth=mo.DimMillimeter(self.machine_variant.l_st)),
            )

        self.comp_winding_layer2 = mo.Component(
            name="WindingLayer2",
            cross_sections=[self.winding_layer2],
            material=mo.MaterialGeneric(name=self.machine_variant.coil_mat["coil_material"]),
            make_solid=mo.MakeExtrude(location=mo.Location3D(), 
            dim_depth=mo.DimMillimeter(self.machine_variant.l_st)),
            )

        self.comp_rotor_core = mo.Component(
            name="RotorCore",
            cross_sections=[self.rotor_core],
            material=mo.MaterialGeneric(name=self.machine_variant.rotor_iron_mat["core_material"], color=r"#808080"),
            make_solid=mo.MakeExtrude(location=mo.Location3D(), 
                    dim_depth=mo.DimMillimeter(self.machine_variant.l_st)),
            )


        tool.bMirror = False

        tool.sketch = tool.create_sketch()
        tool.sketch.SetProperty("Name", self.stator_core.name)
        tool.sketch.SetProperty("Color", r"#808080")
        cs_stator = self.stator_core.draw(tool)
        stator_tool = tool.prepare_section(cs_stator, self.machine_variant.Q)

        tool.sketch = tool.create_sketch()
        tool.sketch.SetProperty("Name", self.winding_layer1.name)
        tool.sketch.SetProperty("Color", r"#B87333")
        cs_winding_layer1 = self.winding_layer1.draw(tool)
        winding_tool1 = tool.prepare_section(cs_winding_layer1, self.machine_variant.Q)
        self.winding_layer1_inner_coord = cs_winding_layer1.inner_coord

        tool.sketch = tool.create_sketch()
        tool.sketch.SetProperty("Name", self.winding_layer2.name)
        tool.sketch.SetProperty("Color", r"#B87333")
        cs_winding_layer2 = self.winding_layer2.draw(tool)
        winding_tool2 = tool.prepare_section(cs_winding_layer2, self.machine_variant.Q)
        self.winding_layer2_inner_coord = cs_winding_layer2.inner_coord

        tool.sketch = tool.create_sketch()
        tool.sketch.SetProperty("Name", self.rotor_core.name)
        tool.sketch.SetProperty("Color", r"#808080")
        cs_rotor_core = self.rotor_core.draw(tool)
        rotor_tool = tool.prepare_section(cs_rotor_core, self.machine_variant.p/self.machine_variant.p)

        tool.sketch = tool.create_sketch()
        tool.sketch.SetProperty("Name", self.shaft.name)
        tool.sketch.SetProperty("Color", r"#71797E")
        cs_shaft = self.shaft.draw(tool)
        shaft_tool = tool.prepare_section(cs_shaft)

        return True

    def show(self, name, toString=False):
        attrs = list(vars(self).items())
        key_list = [el[0] for el in attrs]
        val_list = [el[1] for el in attrs]
        the_dict = dict(list(zip(key_list, val_list)))
        sorted_key = sorted(
            key_list,
            key=lambda item: (
                int(item.partition(" ")[0]) if item[0].isdigit() else float("inf"),
                item,
            ),
        )  # this is also useful for string beginning with digiterations '15 Steel'.
        tuple_list = [(key, the_dict[key]) for key in sorted_key]
        if not toString:
            print("- Bearingless BIM Individual #%s\n\t" % name, end=" ")
            print(", \n\t".join("%s = %s" % item for item in tuple_list))
            return ""
        else:
            return "\n- Bearingless BIM Individual #%s\n\t" % name + ", \n\t".join(
                "%s = %s" % item for item in tuple_list
            )

    def pre_process(self, model):
        # pre-process : you can select part by coordinate!
        """Group"""

        def group(name, id_list):
            model.GetGroupList().CreateGroup(name)
            for the_id in id_list:
                model.GetGroupList().AddPartToGroup(name, the_id)
                # model.GetGroupList().AddPartToGroup(name, name) #<- this also works

        part_ID_list = model.GetPartIDs()

        if len(part_ID_list) != int(
            1 + 1 + self.machine_variant.Q * 2 + 1
        ):
            print("Parts are missing in this machine")
            return False

        self.id_statorCore = id_statorCore = part_ID_list[0]
        partIDRange_Coil = part_ID_list[1 : int(2 * self.machine_variant.Q + 1)]
        self.id_rotorCore = id_rotorCore = part_ID_list[int(2 * self.machine_variant.Q + 1)]
        id_shaft = part_ID_list[-1]

        group("Coils", partIDRange_Coil)

        """ Add Part to Set for later references """

        def add_part_to_set(name, x, y, ID=None):
            model.GetSetList().CreatePartSet(name)
            model.GetSetList().GetSet(name).SetMatcherType("Selection")
            model.GetSetList().GetSet(name).ClearParts()
            sel = model.GetSetList().GetSet(name).GetSelection()
            if ID is None:
                # print x,y
                sel.SelectPartByPosition(x, y, 0)  # z=0 for 2D
            else:
                sel.SelectPart(ID)
            model.GetSetList().GetSet(name).AddSelected(sel)

        # Shaft
        add_part_to_set("ShaftSet", 0.0, 0.0, ID=id_shaft)

        # Create Set for right layer
        Angle_StatorSlotSpan = 360 / self.machine_variant.Q
        # R = self.r_si + self.d_sp + self.d_st *0.5 # this is not generally working (JMAG selects stator core instead.)
        # THETA = 0.25*(Angle_StatorSlotSpan)/180.*np.pi
        R = np.sqrt(self.winding_layer1_inner_coord[0] ** 2 + self.winding_layer1_inner_coord[1] ** 2)
        THETA = np.arctan(self.winding_layer1_inner_coord[1] / self.winding_layer1_inner_coord[0])
        X = R * np.cos(THETA)
        Y = R * np.sin(THETA)
        count = 0
        for UVW, UpDown in zip(
            self.machine_variant.layer_phases[0], self.machine_variant.layer_polarity[0]
        ):
            count += 1
            add_part_to_set("coil_right_%s%s %d" % (UVW, UpDown, count), X, Y)

            # print(X, Y, THETA)
            THETA += Angle_StatorSlotSpan / 180.0 * np.pi
            X = R * np.cos(THETA)
            Y = R * np.sin(THETA)

        # Create Set for left layer
        THETA = np.arctan(self.winding_layer2_inner_coord[1] / self.winding_layer2_inner_coord[0])
        X = R * np.cos(THETA)
        Y = R * np.sin(THETA)
        count = 0
        for UVW, UpDown in zip(
            self.machine_variant.layer_phases[1], self.machine_variant.layer_polarity[1]
        ):
            count += 1
            add_part_to_set("coil_left_%s%s %d" % (UVW, UpDown, count), X, Y)

            THETA += Angle_StatorSlotSpan / 180.0 * np.pi
            X = R * np.cos(THETA)
            Y = R * np.sin(THETA)

        # Create Set for Motion Region
        def part_list_set(name, list_part_id=None, prefix=None):
            model.GetSetList().CreatePartSet(name)
            model.GetSetList().GetSet(name).SetMatcherType("Selection")
            model.GetSetList().GetSet(name).ClearParts()
            sel = model.GetSetList().GetSet(name).GetSelection()
            if list_part_id is not None:
                for ID in list_part_id:
                    sel.SelectPart(ID)
            model.GetSetList().GetSet(name).AddSelected(sel)

        part_list_set(
            "Motion_Region", list_part_id=[id_rotorCore, id_shaft]
        )

        return True

    def create_custom_material(self, app, steel_name):

        core_mat_obj = app.GetMaterialLibrary().GetCustomMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        )
        app.GetMaterialLibrary().DeleteCustomMaterialByObject(core_mat_obj)

        app.GetMaterialLibrary().CreateCustomMaterial(
            self.machine_variant.stator_iron_mat["core_material"], "Custom Materials"
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "Density", self.machine_variant.stator_iron_mat["core_material_density"]
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("MagneticSteelPermeabilityType", 2)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("CoerciveForce", 0)
        # app.GetMaterialLibrary().GetUserMaterial(u"Arnon5-final").GetTable("BhTable").SetName(u"SmoothZeroPointOne")
        BH = np.loadtxt(
            self.machine_variant.stator_iron_mat["core_bh_file"],
            unpack=True,
            usecols=(0, 1),
        )  # values from Nishanth Magnet BH curve
        refarray = BH.T.tolist()
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).GetTable("BhTable").SetTable(refarray)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("DemagnetizationCoerciveForce", 0)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("MagnetizationSaturated", 0)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("MagnetizationSaturated2", 0)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("ExtrapolationMethod", 1)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "YoungModulus", self.machine_variant.stator_iron_mat["core_youngs_modulus"]
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue("Loss_Type", 1)
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "LossConstantKhX", self.machine_variant.stator_iron_mat["core_ironloss_Kh"]
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "LossConstantKeX", self.machine_variant.stator_iron_mat["core_ironloss_Ke"]
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "LossConstantAlphaX",
            self.machine_variant.stator_iron_mat["core_ironloss_a"],
        )
        app.GetMaterialLibrary().GetUserMaterial(
            self.machine_variant.stator_iron_mat["core_material"]
        ).SetValue(
            "LossConstantBetaX", self.machine_variant.stator_iron_mat["core_ironloss_b"]
        )

    def add_transient_magnetic_study(
        self, app, model, dir_csv_output_folder, study_name
    ):

        # study = toolJmag.create_study(self.study_name, "Transient2D", model)
        model.CreateStudy("Transient2D", study_name)
        app.SetCurrentStudy(self.study_name)
        study = model.GetStudy(study_name)

        # Study properties
        study.GetStudyProperties().SetValue("ConversionType", 0)
        study.GetStudyProperties().SetValue(
            "NonlinearMaxIteration", self.config.max_nonlinear_iterations
        )
        study.GetStudyProperties().SetValue(
            "ModelThickness", self.machine_variant.l_st
        )  # [mm] Stack Length

        # Material
        self.add_materials(study)

        # Conditions - Motion
        self.the_speed = self.drive_freq * 60.0 / self.machine_variant.p  # rpm
        study.CreateCondition(
            "RotationMotion", "RotCon"
        )  # study.GetCondition(u"RotCon").SetXYZPoint(u"", 0, 0, 1) # megbox warning
        print("Speed in RPM", self.the_speed)
        study.GetCondition("RotCon").SetValue("AngularVelocity", int(self.the_speed))
        study.GetCondition("RotCon").ClearParts()
        study.GetCondition("RotCon").AddSet(
            model.GetSetList().GetSet("Motion_Region"), 0
        )

        # Implementation of id=0 control:
        #   d-axis initial position is self.alpha_m*0.5
        #   The U-phase current is sin(omega_syn*t) = 0 at t=0.
        #study.GetCondition("RotCon").SetValue(
        #    "InitialRotationAngle",
        #    -self.machine_variant.alpha_m * 0.5
        #    + 90
        #    + self.initial_excitation_bias_compensation_deg()
        #    + (180 / self.machine_variant.p),
        #)

        study.CreateCondition(
            "Torque", "TorCon"
        )  # study.GetCondition(u"TorCon").SetXYZPoint(u"", 0, 0, 0) # megbox warning
        study.GetCondition("TorCon").SetValue("TargetType", 1)
        study.GetCondition("TorCon").SetLinkWithType("LinkedMotion", "RotCon")
        study.GetCondition("TorCon").ClearParts()

        study.CreateCondition("Force", "ForCon")
        study.GetCondition("ForCon").SetValue("TargetType", 1)
        study.GetCondition("ForCon").SetLinkWithType("LinkedMotion", "RotCon")
        study.GetCondition("ForCon").ClearParts()

        # Conditions - FEM Coils & Conductors (i.e. stator/rotor winding)
        self.add_circuit(app, model, study, bool_3PhaseCurrentSource=False)

        # True: no mesh or field results are needed
        study.GetStudyProperties().SetValue(
            "OnlyTableResults", self.config.only_table_results
        )

        # this can be said to be super fast over ICCG solver.
        # https://www2.jmag-international.com/support/en/pdf/JMAG-Designer_Ver.17.1_ENv3.pdf
        study.GetStudyProperties().SetValue("DirectSolverType", 1)

        if self.config.multiple_cpus:
            # This SMP(shared memory process) is effective only if there are tons of elements. e.g., over 100,000.
            # too many threads will in turn make them compete with each other and slow down the solve. 2 is good enough
            # for eddy current solve. 6~8 is enough for transient solve.
            study.GetStudyProperties().SetValue("UseMultiCPU", True)
            study.GetStudyProperties().SetValue("MultiCPU", self.config.num_cpus)

        # two sections of different time step
        number_of_revolution_1TS = self.config.no_of_rev_1TS
        number_of_revolution_2TS = self.config.no_of_rev_2TS
        number_of_steps_1TS = (
            self.config.no_of_steps_per_rev_1TS * number_of_revolution_1TS
        )
        number_of_steps_2TS = (
            self.config.no_of_steps_per_rev_2TS * number_of_revolution_2TS
        )
        DM = app.GetDataManager()
        DM.CreatePointArray("point_array/timevsdivision", "SectionStepTable")
        refarray = [[0 for i in range(3)] for j in range(3)]
        refarray[0][0] = 0
        refarray[0][1] = 1
        refarray[0][2] = 50
        refarray[1][0] = number_of_revolution_1TS / self.drive_freq
        refarray[1][1] = number_of_steps_1TS
        refarray[1][2] = 50
        refarray[2][0] = (
            number_of_revolution_1TS + number_of_revolution_2TS
        ) / self.drive_freq
        refarray[2][1] = number_of_steps_2TS  # number_of_steps_2TS
        refarray[2][2] = 50
        DM.GetDataSet("SectionStepTable").SetTable(refarray)
        number_of_total_steps = (
            1 + number_of_steps_1TS + number_of_steps_2TS
        )  # don't forget to modify here!
        study.GetStep().SetValue("Step", number_of_total_steps)
        study.GetStep().SetValue("StepType", 3)
        study.GetStep().SetTableProperty("Division", DM.GetDataSet("SectionStepTable"))

        # add equations
        study.GetDesignTable().AddEquation("freq")
        study.GetDesignTable().AddEquation("speed")
        study.GetDesignTable().GetEquation("freq").SetType(0)
        study.GetDesignTable().GetEquation("freq").SetExpression(
            "%g" % self.drive_freq
        )
        study.GetDesignTable().GetEquation("freq").SetDescription(
            "Excitation Frequency"
        )

        study.GetDesignTable().GetEquation("speed").SetType(1)
        study.GetDesignTable().GetEquation("speed").SetExpression(
            "freq * %d" % (60 / self.machine_variant.p)
        )
        study.GetDesignTable().GetEquation("speed").SetDescription(
            "mechanical speed of four pole"
        )

        # speed, freq, slip
        study.GetCondition("RotCon").SetValue("AngularVelocity", "speed")

        # Iron Loss Calculation Condition
        # Stator
        if True:
            cond = study.CreateCondition("Ironloss", "IronLossConStator")
            cond.SetValue("RevolutionSpeed", "freq*60/%d" % self.machine_variant.p)
            cond.ClearParts()
            sel = cond.GetSelection()
            EPS = 1e-2  # unit: mm
            sel.SelectPartByPosition(self.machine_variant.r_si * 1e3 + EPS, EPS, 0)
            cond.AddSelected(sel)
            # Use FFT for hysteresis to be consistent with FEMM's results and to have a FFT plot
            cond.SetValue("HysteresisLossCalcType", 1)
            cond.SetValue("PresetType", 3)  # 3:Custom
            # Specify the reference steps yourself because you don't really know what JMAG is doing behind you
            cond.SetValue(
                "StartReferenceStep",
                number_of_total_steps + 1 - number_of_steps_2TS * 0.5,
            )  # 1/4 period = number_of_steps_2TS*0.5
            cond.SetValue("EndReferenceStep", number_of_total_steps)
            cond.SetValue("UseStartReferenceStep", 1)
            cond.SetValue("UseEndReferenceStep", 1)
            cond.SetValue(
                "Cyclicity", 4
            )  # specify reference steps for 1/4 period and extend it to whole period
            cond.SetValue("UseFrequencyOrder", 1)
            cond.SetValue("FrequencyOrder", "1-50")  # Harmonics up to 50th orders
        # Check CSV results for iron loss (You cannot check this for Freq study) # CSV and save space
        study.GetStudyProperties().SetValue(
            "CsvOutputPath", dir_csv_output_folder
        )  # it's folder rather than file!
        study.GetStudyProperties().SetValue("CsvResultTypes", self.config.csv_results)
        study.GetStudyProperties().SetValue(
            "DeleteResultFiles", self.config.del_results_after_calc
        )

        # Rotor
        if True:
            cond = study.CreateCondition("Ironloss", "IronLossConRotor")
            cond.SetValue("BasicFrequencyType", 2)
            cond.SetValue("BasicFrequency", "freq")
            # cond.SetValue(u"BasicFrequency", u"slip*freq") # this require the signal length to be at least 1/4 of
            # slip period, that's too long!
            cond.ClearParts()
            sel = cond.GetSelection()
            sel.SelectPart(self.id_rotorCore)

            cond.AddSelected(sel)
            # Use FFT for hysteresis to be consistent with FEMM's results
            cond.SetValue("HysteresisLossCalcType", 1)
            cond.SetValue("PresetType", 3)
            # Specify the reference steps yourself because you don't really know what JMAG is doing behind you
            cond.SetValue(
                "StartReferenceStep",
                number_of_total_steps + 1 - number_of_steps_2TS * 0.5,
            )  # 1/4 period = number_of_steps_2TS*0.5
            cond.SetValue("EndReferenceStep", number_of_total_steps)
            cond.SetValue("UseStartReferenceStep", 1)
            cond.SetValue("UseEndReferenceStep", 1)
            cond.SetValue(
                "Cyclicity", 4
            )  # specify reference steps for 1/4 period and extend it to whole period
            cond.SetValue("UseFrequencyOrder", 1)
            cond.SetValue("FrequencyOrder", "1-50")  # Harmonics up to 50th orders
        self.study_name = study_name
        return study


    def add_materials(self, study):
        # if 'M19' in self.machine_variant.stator_iron_mat["core_material"]:
        # study.SetMaterialByName(self.comp_stator_core.name, "M-19 Steel Gauge-29")
        # study.GetMaterial(self.comp_stator_core.name).SetValue("Laminated", 1)
        # study.GetMaterial(self.comp_stator_core.name).SetValue("LaminationFactor",
        #     self.machine_variant.stator_iron_mat["core_stacking_factor"])

        # study.SetMaterialByName(self.comp_rotor_core.name, "M-19 Steel Gauge-29")
        # study.GetMaterial(self.comp_rotor_core.name).SetValue("Laminated", 1)
        # study.GetMaterial(self.comp_rotor_core.name).SetValue("LaminationFactor",
        #     self.machine_variant.rotor_iron_mat["core_stacking_factor"])

        study.SetMaterialByName(self.comp_stator_core.name,
            self.machine_variant.stator_iron_mat["core_material"])
        study.GetMaterial(self.comp_stator_core.name).SetValue("Laminated", 1)
        study.GetMaterial(self.comp_stator_core.name).SetValue("LaminationFactor",
            self.machine_variant.stator_iron_mat["core_stacking_factor"])

        study.SetMaterialByName(self.comp_rotor_core.name, 
            self.machine_variant.rotor_iron_mat["core_material"])
        study.GetMaterial(self.comp_rotor_core.name).SetValue("Laminated", 1)
        study.GetMaterial(self.comp_rotor_core.name).SetValue("LaminationFactor",
            self.machine_variant.rotor_iron_mat["core_stacking_factor"])

        study.SetMaterialByName(
            "Shaft", self.machine_variant.shaft_mat["shaft_material"]
        )
        study.GetMaterial("Shaft").SetValue("Laminated", 0)
        study.GetMaterial("Shaft").SetValue("EddyCurrentCalculation", 1)

        study.SetMaterialByName("Coils", "Copper")
        study.GetMaterial("Coils").SetValue("UserConductivityType", 1)


    def add_circuit(self, app, model, study, bool_3PhaseCurrentSource=True):
        def add_mp_circuit(study, turns, Rs, x=10, y=10):
            # Placing coils/phase windings
            coil_name = []
            for i in range(0, self.machine_variant.no_of_phases):
                coil_name.append("coil_" + self.machine_variant.name_phases[i])
                study.GetCircuit().CreateComponent("Coil", 
                    coil_name[i])
                study.GetCircuit().CreateInstance(coil_name[i],
                    x + 4 * i, y)
                study.GetCircuit().GetComponent(coil_name[i]).SetValue("Turn", turns)
                study.GetCircuit().GetComponent(coil_name[i]).SetValue("Resistance", Rs)
                study.GetCircuit().GetInstance(coil_name[i], 0).RotateTo(90)

            self.coil_name = coil_name

            # Connecting all phase windings to a neutral point
            for i in range(0, self.machine_variant.no_of_phases - 1):         
                study.GetCircuit().CreateWire(x + 4 * i, y - 2, x + 4 * (i + 1), y - 2)

            study.GetCircuit().CreateComponent("Ground", "Ground")
            study.GetCircuit().CreateInstance("Ground", x + 8, y - 4)
            
            # Placing current sources
            cs_name = []
            for i in range(0, self.machine_variant.no_of_phases):
                cs_name.append("cs_" + self.machine_variant.name_phases[i])
                study.GetCircuit().CreateComponent("CurrentSource", cs_name[i])
                study.GetCircuit().CreateInstance(cs_name[i], x + 4 * i, y + 4)
                study.GetCircuit().GetInstance(cs_name[i], 0).RotateTo(90)

            self.cs_name = cs_name

            # Terminal Voltage/Circuit Voltage: Check for outputting CSV results
            terminal_name = []
            for i in range(0, self.machine_variant.no_of_phases):
                terminal_name.append("vp_" + self.machine_variant.name_phases[i])
                study.GetCircuit().CreateTerminalLabel(terminal_name[i], x + 4 * i, y + 2)
                study.GetCircuit().CreateComponent("VoltageProbe", terminal_name[i])
                study.GetCircuit().CreateInstance(terminal_name[i], x + 2 + 4 * i, y + 2)
                study.GetCircuit().GetInstance(terminal_name[i], 0).RotateTo(90)

            self.terminal_name = terminal_name

        app.ShowCircuitGrid(True)
        study.CreateCircuit()

        add_mp_circuit(study, self.machine_variant.Z_q, Rs=self.R_wdg)

        for phase_name in self.machine_variant.name_phases:
            study.CreateCondition("FEMCoil", phase_name)
            # link between FEM Coil Condition and Circuit FEM Coil
            condition = study.GetCondition(phase_name)
            condition.SetLink("coil_%s" % (phase_name))
            condition.GetSubCondition("untitled").SetName("delete")

        count = 0  # count indicates which slot the current rightlayer is in.
        index = 0
        dict_dir = {"+": 1, "-": 0}
        coil_pitch = self.machine_variant.pitch  # self.dict_coil_connection[0]
        # select the part (via `Set') to assign the FEM Coil condition
        for UVW, UpDown in zip(
            self.machine_variant.layer_phases[0], self.machine_variant.layer_polarity[0]
        ):

            count += 1
            condition = study.GetCondition(UVW)

            # right layer
            # print (count, "Coil Set %d"%(count), end=' ')
            condition.CreateSubCondition("FEMCoilData", "Coil Set Right %d" % count)
            subcondition = condition.GetSubCondition("Coil Set Right %d" % count)
            subcondition.ClearParts()
            subcondition.AddSet(
                model.GetSetList().GetSet(
                    "coil_%s%s%s %d" % ("right_", UVW, UpDown, count)
                ),
                0,
            )  # right layer
            subcondition.SetValue("Direction2D", dict_dir[UpDown])

            # left layer
            if coil_pitch > 0:
                if count + coil_pitch <= self.machine_variant.Q:
                    count_leftlayer = count + coil_pitch
                    index_leftlayer = index + coil_pitch
                else:
                    count_leftlayer = int(count + coil_pitch - self.machine_variant.Q)
                    index_leftlayer = int(index + coil_pitch - self.machine_variant.Q)
            else:
                if count + coil_pitch > 0:
                    count_leftlayer = count + coil_pitch
                    index_leftlayer = index + coil_pitch
                else:
                    count_leftlayer = int(count + coil_pitch + self.machine_variant.Q)
                    index_leftlayer = int(index + coil_pitch + self.machine_variant.Q)

            # Check if it is a distributed windg???
            if self.machine_variant.pitch == 1:
                UVW = self.machine_variant.layer_phases[1][index_leftlayer]
                UpDown = self.machine_variant.layer_polarity[1][index_leftlayer]
            else:
                if self.machine_variant.layer_phases[1][index_leftlayer] != UVW:
                    print("[Warn] Potential bug in your winding layout detected.")
                    raise Exception("Bug in winding layout detected.")
                if UpDown == "+":
                    UpDown = "-"
                else:
                    UpDown = "+"
            # print (count_leftlayer, "Coil Set %d"%(count_leftlayer))
            condition.CreateSubCondition(
                "FEMCoilData", "Coil Set Left %d" % count_leftlayer
            )
            subcondition = condition.GetSubCondition(
                "Coil Set Left %d" % count_leftlayer
            )
            subcondition.ClearParts()
            subcondition.AddSet(
                model.GetSetList().GetSet(
                    "coil_%s%s%s %d" % ("left_", UVW, UpDown, count_leftlayer)
                ),
                0,
            )  # left layer
            subcondition.SetValue("Direction2D", dict_dir[UpDown])
            index += 1
            # clean up
            for phase_name in self.machine_variant.name_phases:
                condition = study.GetCondition(phase_name)
                condition.RemoveSubCondition("delete")


    def set_currents_two_sequences(self, Is1, Is2, s1, s2, freq, phi_s1_0, phi_s2_0, app, study):
        # Setting current values after creating a circuit using "add_mp_circuit" method
        # "freq" variable cannot be used here. So pay extra attention when you 
        # create new case of a different freq.
        for i in range(0, self.machine_variant.no_of_phases):
            func = app.FunctionFactory().Composite()
            f1 = app.FunctionFactory().Sin(Is1, freq,
                - s1 * 360 / self.machine_variant.no_of_phases * i + phi_s1_0 + 90)
            f2 = app.FunctionFactory().Sin(Is2, freq,
                - s2 * 360 / self.machine_variant.no_of_phases * i + phi_s2_0 + 90)
            func.AddFunction(f1)
            func.AddFunction(f2)
            study.GetCircuit().GetComponent(self.cs_name[i]).SetFunction(func)

    def set_currents_multiple_sequences(self, seq_orders, seq_freqs, seq_ampls,
         seq_phase_shifts, app, study):
        m = self.machine_variant.no_of_phases
        for i in range(0, m):
            func = app.FunctionFactory().Composite()
            for j in range(0, len(seq_orders)):
                if seq_freqs[j] == 'drive_freq':
                    f = app.FunctionFactory().Sin(seq_ampls[j], self.drive_freq,
                        - seq_orders[j] * 360 / m * i + seq_phase_shifts[j] + 90)
                elif seq_freqs[j] == 0:
                    f = app.FunctionFactory().Constant(
                        seq_ampls[j] * np.cos(
                            (- seq_orders[j] * 360 / m * i + seq_phase_shifts[j]) / 180 * np.pi
                            )
                        )
                else:
                    f = app.FunctionFactory().Sin(seq_ampls[j], seq_freqs[j],
                        - seq_orders[j] * 360 / m * i + seq_phase_shifts[j] + 90)
                
                func.AddFunction(f)

            study.GetCircuit().GetComponent(self.cs_name[i]).SetFunction(func)

    def set_currents_four_sequences(self, Is1, Is2, Is3, Is4, s1, s2, s3, s4,
         freq, phi_s1_0, phi_s2_0, phi_s3_0, phi_s4_0, app, study):
        for i in range(0, self.machine_variant.no_of_phases):
            func = app.FunctionFactory().Composite()
            f1 = app.FunctionFactory().Sin(Is1, freq,
                - s1 * 360 / self.machine_variant.no_of_phases * i + phi_s1_0 + 90)
            f2 = app.FunctionFactory().Sin(Is2, freq,
                - s2 * 360 / self.machine_variant.no_of_phases * i + phi_s2_0 + 90)
            f3 = app.FunctionFactory().Sin(Is3, freq,
                - s3 * 360 / self.machine_variant.no_of_phases * i + phi_s3_0 + 90)
            f4 = app.FunctionFactory().Sin(Is4, freq,
                - s4 * 360 / self.machine_variant.no_of_phases * i + phi_s4_0 + 90)
            func.AddFunction(f1)
            func.AddFunction(f2)
            func.AddFunction(f3)
            func.AddFunction(f4)
            study.GetCircuit().GetComponent(self.cs_name[i]).SetFunction(func)


    def add_time_step_settings(self, time1_interval, no_of_steps_1st_TSS,
     time2_interval, no_of_steps_2nd_TSS, app, study):

        DM = app.GetDataManager()
        DM.CreatePointArray("point_array/timevsdivision", "SectionStepTable")
        refarray = [[0 for i in range(3)] for j in range(3)]
        refarray[0][0] = 0
        refarray[0][1] = 1
        refarray[0][2] = 50
        refarray[1][0] = time1_interval
        refarray[1][1] = no_of_steps_1st_TSS
        refarray[1][2] = 50
        refarray[2][0] = refarray[1][0] + time2_interval
        refarray[2][1] = no_of_steps_2nd_TSS
        refarray[2][2] = 50
        DM.GetDataSet("SectionStepTable").SetTable(refarray)
        number_of_total_steps = (
            1 + no_of_steps_1st_TSS + no_of_steps_2nd_TSS
        )  # don't forget to modify here!
        study.GetStep().SetValue("Step", number_of_total_steps)
        study.GetStep().SetValue("StepType", 3)
        study.GetStep().SetTableProperty("Division", DM.GetDataSet("SectionStepTable"))


    def mesh_study(self, app, model, study):

        # this `if' judgment is effective only if JMAG-DeleteResultFiles is False
        # if not study.AnyCaseHasResult():
        # mesh
        print("------------------Adding mesh")
        self.add_mesh(study, model)

        # Export Image
        app.View().ShowAllAirRegions()
        # app.View().ShowMeshGeometry() # 2nd btn
        app.View().ShowMesh()  # 3rn btn
        app.View().Zoom(3)
        app.View().Pan(-self.machine_variant.r_si / 1000, 0)
        app.ExportImageWithSize(
            self.design_results_folder + self.project_name + "mesh.png", 2000, 2000
        )
        app.View().ShowModel()  # 1st btn. close mesh view, and note that mesh data will be deleted if only ouput table
        # results are selected.


    def add_mesh(self, study, model):
        # this is for multi slide planes, which we will not be usin
        refarray = [[0 for i in range(2)] for j in range(1)]
        refarray[0][0] = 3
        refarray[0][1] = 1
        study.GetMeshControl().GetTable("SlideTable2D").SetTable(refarray)

        study.GetMeshControl().SetValue("MeshType", 1)  # make sure this has been exe'd:
        # study.GetCondition(u"RotCon").AddSet(model.GetSetList().GetSet(u"Motion_Region"), 0)
        study.GetMeshControl().SetValue(
            "RadialDivision", self.config.airgap_mesh_radial_div
        )  # for air region near which motion occurs
        study.GetMeshControl().SetValue(
            "CircumferentialDivision", self.config.airgap_mesh_circum_div
        )  # 1440) # for air region near which motion occurs
        study.GetMeshControl().SetValue(
            "AirRegionScale", self.config.mesh_air_region_scale
        )  # [Model Length]: Specify a value within (1.05 <= value < 1000)
        study.GetMeshControl().SetValue("MeshSize", self.config.mesh_size)
        study.GetMeshControl().SetValue("AutoAirMeshSize", 0)
        study.GetMeshControl().SetValue(
            "AirMeshSize", self.config.mesh_size
        )  # mm
        study.GetMeshControl().SetValue("Adaptive", 0)

        study.GetMeshControl().CreateCondition("RotationPeriodicMeshAutomatic", 
                "autoRotMesh") # with this you can choose to set CircumferentialDivision automatically

        study.GetMeshControl().CreateCondition("Part", "RotorMeshCtrl")
        study.GetMeshControl().GetCondition("RotorMeshCtrl").SetValue("Size", self.config.mesh_size_rotor)
        study.GetMeshControl().GetCondition("RotorMeshCtrl").ClearParts()
        study.GetMeshControl().GetCondition("RotorMeshCtrl").AddSet(model.GetSetList().GetSet("RotorSet"), 0)

        study.GetMeshControl().CreateCondition("Part", "ShaftMeshCtrl")
        study.GetMeshControl().GetCondition("ShaftMeshCtrl").SetValue("Size", 10) # 10 mm
        study.GetMeshControl().GetCondition("ShaftMeshCtrl").ClearParts()
        study.GetMeshControl().GetCondition("ShaftMeshCtrl").AddSet(model.GetSetList().GetSet("ShaftSet"), 0)

        def mesh_all_cases(study):
            numCase = study.GetDesignTable().NumCases()
            for case in range(0, numCase):
                study.SetCurrentCase(case)
                if not study.HasMesh():
                    study.CreateMesh()

        # if self.MODEL_ROTATE:
        #     if self.total_number_of_cases>1: # just to make sure
        #         model.RestoreCadLink()
        #         study.ApplyAllCasesCadParameters()

        mesh_all_cases(study)


    def run_study(self, app, study, toc):
        if not self.config.jmag_scheduler:
            print("-----------------------Running JMAG...")
            # if run_list[1] == True:
            study.RunAllCases()
            msg = "Time spent on %s is %g s." % (study.GetName(), clock_time() - toc)
            print(msg)
        else:
            print("Submit to JMAG_Scheduler...")
            job = study.CreateJob()
            job.SetValue("Title", study.GetName())
            job.SetValue("Queued", True)
            job.Submit(False)  # False:CurrentCase, True:AllCases
            # wait and check
            # study.CheckForCaseResults()
        
        app.Save()


    def prepare_section(
        self, list_regions, tool, bMirrorMerge=True, bRotateMerge=True
    ):  # csToken is a list of cross section's token

        def regionCircularPattern360Origin(region, tool, bMerge=True):
            # index is used to define name of region

            Q_float = float(tool.iRotateCopy)  # don't ask me, ask JSOL
            circular_pattern = tool.sketch.CreateRegionCircularPattern()
            circular_pattern.SetProperty("Merge", bMerge)

            ref2 = tool.doc.CreateReferenceFromItem(region)
            circular_pattern.SetPropertyByReference("Region", ref2)
            face_region_string = circular_pattern.GetProperty("Region")

            circular_pattern.SetProperty("CenterType", 2)  # origin I guess

            # print('Copy', Q_float)
            circular_pattern.SetProperty("Angle", "360/%d" % Q_float)
            circular_pattern.SetProperty("Instance", str(Q_float))

        list_region_objects = []
        for idx, list_segments in enumerate(list_regions):
            # Region
            tool.doc.GetSelection().Clear()
            for segment in list_segments:
                tool.doc.GetSelection().Add(tool.sketch.GetItem(segment.draw_token.GetName()))

            tool.sketch.CreateRegions()
            # self.sketch.CreateRegionsWithCleanup(EPS, True) # StatorCore will fail

            if idx == 0:
                region_object = tool.sketch.GetItem(
                    "Region"
                )  # This is how you get access to the region you create.
            else:
                region_object = tool.sketch.GetItem(
                    "Region.%d" % (idx + 1)
                )  # This is how you get access to the region you create.
            list_region_objects.append(region_object)
        # raise

        for idx, region_object in enumerate(list_region_objects):
            # Mirror
            if tool.bMirror == True:
                if tool.edge4Ref is None:
                    tool.regionMirrorCopy(
                        region_object,
                        edge4Ref=None,
                        symmetryType=2,
                        bMerge=bMirrorMerge,
                    )  # symmetryType=2 means x-axis as ref
                else:
                    tool.regionMirrorCopy(
                        region_object,
                        edge4Ref=tool.edge4ref,
                        symmetryType=None,
                        bMerge=bMirrorMerge,
                    )  # symmetryType=2 means x-axis as ref

            # RotateCopy
            if tool.iRotateCopy != 0:
                # print('Copy', self.iRotateCopy)
                regionCircularPattern360Origin(
                    region_object, tool, bMerge=bRotateMerge
                )

        tool.sketch.CloseSketch()
        return list_region_objects

    def extract_JMAG_results(self, path, study_name):
        current_csv_path = path + study_name + "_circuit_current.csv"
        voltage_csv_path = path + study_name + "_circuit_voltage.csv"
        torque_csv_path = path + study_name + "_torque.csv"
        force_csv_path = path + study_name + "_force.csv"
        iron_loss_path = path + study_name + "_iron_loss_loss.csv"
        hysteresis_loss_path = path + study_name + "_hysteresis_loss_loss.csv"
        eddy_current_loss_path = path + study_name + "_joule_loss_loss.csv"
        ohmic_loss_path = path + study_name + "_joule_loss.csv"

        curr_df = pd.read_csv(current_csv_path, skiprows=7)
        volt_df = pd.read_csv(voltage_csv_path, skiprows=7)
        tor_df = pd.read_csv(torque_csv_path, skiprows=7)
        force_df = pd.read_csv(force_csv_path, skiprows=7)
        iron_df = pd.read_csv(iron_loss_path, skiprows=7)
        hyst_df = pd.read_csv(hysteresis_loss_path, skiprows=7)
        eddy_df = pd.read_csv(eddy_current_loss_path, skiprows=7)
        ohmic_df = pd.read_csv(ohmic_loss_path, skiprows=7)

        curr_df = curr_df.set_index("Time(s)")
        volt_df = volt_df.set_index("Time(s)")
        tor_df = tor_df.set_index("Time(s)")
        force_df = force_df.set_index("Time(s)")
        eddy_df = eddy_df.set_index("Frequency(Hz)")
        hyst_df = hyst_df.set_index("Frequency(Hz)")
        iron_df = iron_df.set_index("Frequency(Hz)")
        ohmic_df = ohmic_df.set_index("Time(s)")

        fea_data = {
            "current": curr_df,
            "voltage": volt_df,
            "torque": tor_df,
            "force": force_df,
            "iron_loss": iron_df,
            "hysteresis_loss": hyst_df,
            "eddy_current_loss": eddy_df,
            "ohmic_loss": ohmic_df,
            "no_of_steps_2nd_TSS": self.config.no_of_steps_2nd_TSS,
            "no_of_rev_2nd_TSS": self.config.no_of_rev_2nd_TSS,
            "scale_axial_length": self.config.scale_axial_length,           
            "drive_freq": self.drive_freq,
            "stator_wdg_resistances": [self.R_wdg, self.R_wdg_coil_ends, self.R_wdg_coil_sides],
            "conductor_names": self.conductor_names,
            "breakdown_torque_from_tha": self.breakdown_torque,
            "rotor_current_tha": self.rotor_current_tha,
            "rotor_slot_area_tha": self.rotor_slot_area_tha,
            "stator_slot_area_tha": self.stator_slot_area_tha,
        }

        return fea_data