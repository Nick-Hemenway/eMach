import os
import sys
from copy import deepcopy

import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import pygmo as pg
import numpy as np
import pandas as pd

from datahandler import DataHandler


sys.path.append("..")

path = os.path.abspath('')
arch_file = path + r'\opti_archive_data.pkl'  # specify path where saved data will reside
des_file = path + r'\opti_designer.pkl'
pop_file = path + r'\latest_pop.csv'
dh = DataHandler(arch_file, des_file)  # initialize data handler with required file paths


os.chdir(os.path.dirname(__file__))
sys.path.append("../../../..")

from machine_design import BSPMArchitectType1
from specifications.bspm_specification import BSPMMachineSpec
from specifications.machine_specs.bp1_machine_specs import DesignSpec
from specifications.materials.electric_steels import Arnon5
from specifications.materials.jmag_library_magnets import N40H
from specifications.materials.miscellaneous_materials import (
    CarbonFiber,
    Steel,
    Copper,
    Hub,
    Air,
)
from specifications.analyzer_config.em_fea_config import JMAG_FEA_Configuration

from problems.bspm_em_problem import BSPM_EM_Problem
from post_analyzers.bpsm_em_post_analyzer import BSPM_EM_PostAnalyzer
from length_scale_step import LengthScaleStep

from settings.bspm_settings_handler import BSPM_Settings_Handler
from local_analyzers.em import BSPM_EM_Analysis
from local_analyzers import structrual_analyzer as sta
from local_analyzers import thermal_analyzer as therm

from bspm_ds import BSPMDesignSpace
from mach_eval import AnalysisStep, MachineDesigner, MachineEvaluator, State
from mach_opt import DesignProblem, DesignOptimizationMOEAD, InvalidDesign

from datahandler import DataHandler

##############################################################################
############################ Define Design ###################################
##############################################################################

# create specification object for the BSPM machine
machine_spec = BSPMMachineSpec(
    design_spec=DesignSpec,
    rotor_core=Arnon5,
    stator_core=Arnon5,
    magnet=N40H,
    conductor=Copper,
    shaft=Steel,
    air=Air,
    sleeve=CarbonFiber,
    hub=Hub,
)


##############################################################################
############################ Define Struct AnalysisStep ######################
##############################################################################
stress_limits = {
    "rad_sleeve": -100e6,
    "tan_sleeve": 1300e6,
    "rad_magnets": 0,
    "tan_magnets": 80e6,
}
# spd = sta.SleeveProblemDef(design_variant)
# problem = spd.get_problem()
struct_ana = sta.SleeveAnalyzer(stress_limits)


# sleeve_dim = ana.analyze(problem)
# print(sleeve_dim)


class StructPostAnalyzer:
    """Converts a State into a problem"""

    def get_next_state(results, in_state):
        if results is False:
            raise InvalidDesign("Suitable sleeve not found")
        else:
            print("Results are ", type(results))
            machine = in_state.design.machine
            new_machine = machine.clone(machine_parameter_dict={"d_sl": results[0]})
        state_out = deepcopy(in_state)
        state_out.design.machine = new_machine
        return state_out


struct_step = AnalysisStep(sta.SleeveProblemDef, struct_ana, StructPostAnalyzer)


##############################################################################
############################ Define EM AnalysisStep ##########################
##############################################################################


class BSPM_EM_ProblemDefinition:
    """Converts a State into a problem"""

    def get_problem(state):
        problem = BSPM_EM_Problem(state.design.machine, state.design.settings)
        return problem


# initialize em analyzer class with FEA configuration
em_analysis = BSPM_EM_Analysis(JMAG_FEA_Configuration)

# define em step
em_step = AnalysisStep(BSPM_EM_ProblemDefinition, em_analysis, BSPM_EM_PostAnalyzer)


##############################################################################
############################ Define Thermal AnalysisStep #####################
##############################################################################


class AirflowPostAnalyzer:
    """Converts a State into a problem"""

    def get_next_state(results, in_state):
        if results["valid"] is False:
            raise InvalidDesign("Magnet temperature beyond limits")
        else:
            state_out = deepcopy(in_state)
            state_out.conditions.airflow = results
        return state_out


thermal_step = AnalysisStep(
    therm.AirflowProblemDef, therm.AirflowAnalyzer, AirflowPostAnalyzer
)


##############################################################################
############################ Define Windage AnalysisStep #####################
##############################################################################


class WindageLossPostAnalyzer:
    """Converts a State into a problem"""

    def get_next_state(results, in_state):
        state_out = deepcopy(in_state)
        machine = state_out.design.machine
        eff = (
            100
            * machine.mech_power
            / (
                machine.mech_power
                + results
                + state_out.conditions.em["copper_loss"]
                + state_out.conditions.em["rotor_iron_loss"]
                + state_out.conditions.em["stator_iron_loss"]
                + state_out.conditions.em["magnet_loss"]
            )
        )
        state_out.conditions.windage = {"loss": results, "efficiency": eff}
        return state_out


windage_step = AnalysisStep(
    therm.WindageProblemDef, therm.WindageLossAnalyzer, WindageLossPostAnalyzer
)


#######################################
########## CODE STARTS HERE ###########
#######################################

# Extract data

prt = dh.load_from_archive()
objects = []
variables = []
resultsp = []
designs = []
temp = []

for data in prt:
    objects.append(data.objs)
    variables.append(data.x)
    resultsp.append(data.full_results)
    designs.append(data.design)
    temp = list(data.full_results)
    if temp[1][1]['project_name'] == 'proj_1994_': # design with -ve magnet temperature
        em_conditions = data.full_results[2][0].conditions
        em_design = data.full_results[2][0].design
        
# run thermal analysis

state_in = State(em_design, em_conditions)

problem = thermal_step.problem_definition.get_problem(state_in)
resulting = thermal_step.analyzer.analyze(problem)
state_out = thermal_step.post_analyzer.get_next_state(resulting, state_in)

