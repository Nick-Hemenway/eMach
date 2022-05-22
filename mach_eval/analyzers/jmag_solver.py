from time import time as clock_time
import numpy as np
import pandas as pd

from mach_cad.tools.jmag import JmagDesigner


class JMagTransient2DFEA:
    def run(self, problem: "FEAProblem"):
        """
        This is where the order of operations for the 2D Transient analysis
        is defined. It calls functions implemented in the base class to 
        perform actions and then returns results.
        """
        self.init_tool()
        self.make_components(problem.components)  # will be very slow right now...
        self.apply_materials(problem.components)  # apply the materials

        for cond in problem.conditions:
            cond.apply(self.tool)
        for sett in problem.settings:
            sett.apply(self)
        for get_res in problem.get_results:
            get_res.define(self)
        for config in problem.configs:
            config.apply(self)
        self.run_study()  # run the study

        results = [None] * len(problem.get_results)
        results = self.collect_results(problem.get_results)  # collect results
        self.clear_tool()
        return results

    def init_tool(self):
        self.tool = JmagDesigner()
        self.tool.open(comp_filepath="trial.jproj")
        self.tool.set_visibility(True)

    def make_components(self, components):
        # Draw components
        for component in components:
            component.make(self.tool, self.tool)
        self.tool.save()
        # self.make_sets(components)  # needed for assigning materials later

    # def make_sets(self, components):
    #     magnet_names = []
    #     for component in components:
    #         if "Magnet" in component.name:
    #             magnet_names.append(component.name)

    #     model = self.tool.jd.GetCurrentModel()

    #     def add_part_to_set(name, x, y, ID=None):
    #         x = model.GetSetList().GetSet(name)
    #         if x.IsValid() is False:
    #             model.GetSetList().CreatePartSet(name)
    #         model.GetSetList().GetSet(name).SetMatcherType("Selection")
    #         # model.GetSetList().GetSet(name).ClearParts()
    #         sel = model.GetSetList().GetSet(name).GetSelection()
    #         if ID is None:
    #             # print x,y
    #             sel.SelectPartByPosition(x, y, 0)  # z=0 for 2D
    #         else:
    #             sel.SelectPart(ID)
    #         model.GetSetList().GetSet(name).AddSelected(sel)

    #     def group(name, id_list):
    #         model.GetGroupList().CreateGroup(name)
    #         for the_id in id_list:
    #             model.GetGroupList().AddPartToGroup(name, the_id)
    #             # model.GetGroupList().AddPartToGroup(name, name) #<- this also works

    #     add_part_to_set("Magnet", 0, 0, magnet_names[0])
    #     add_part_to_set("Magnet", 0, 0, magnet_names[1])
    #     return True

    def apply_materials(self, components):
        # update stator and rotor material properties
        self.tool.study.GetMaterial("Stator").SetValue("Laminated", 1)
        self.tool.study.GetMaterial("Stator").SetValue("LaminationFactor", 96)
        self.tool.study.GetMaterial(u"Shaft").SetValue(u"EddyCurrentCalculation", 1)

        # update magnet material properties
        self.tool.study.GetMaterial(u"Magnet0").SetValue(u"EddyCurrentCalculation", 1)
        self.tool.study.GetMaterial(u"Magnet0").SetValue(u"Temperature", 80)
        self.tool.study.GetMaterial(u"Magnet0").SetValue(u"Poles", 2)
        self.tool.study.GetMaterial(u"Magnet0").SetDirectionXYZ(1, 0, 0)
        self.tool.study.GetMaterial(u"Magnet0").SetAxisXYZ(0, 0, 1)
        self.tool.study.GetMaterial(u"Magnet0").SetOriginXYZ(0, 0, 0)
        self.tool.study.GetMaterial(u"Magnet0").SetPattern(u"ParallelCircular")

        self.tool.study.GetMaterial(u"Magnet1").SetValue(u"EddyCurrentCalculation", 1)
        self.tool.study.GetMaterial(u"Magnet1").SetValue(u"Temperature", 80)
        self.tool.study.GetMaterial(u"Magnet1").SetValue(u"Poles", 2)
        self.tool.study.GetMaterial(u"Magnet1").SetDirectionXYZ(1, 0, 0)
        self.tool.study.GetMaterial(u"Magnet1").SetAxisXYZ(0, 0, 1)
        self.tool.study.GetMaterial(u"Magnet1").SetOriginXYZ(0, 0, 0)
        self.tool.study.GetMaterial(u"Magnet1").SetPattern(u"ParallelCircular")

    def run_study(self):
        pass  # run the stufy


class MaterialTemp:
    def __init__(self, name, values):
        self.name = name
        self.values = values  # this is a dictionary of key value pairs for Jmag

