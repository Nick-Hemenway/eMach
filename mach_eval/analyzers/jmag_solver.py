# -*- coding: utf-8 -*-
"""
Created on Thu May  5 15:25:37 2022

@author: Martin Johnson
"""

from time import time as clock_time
import os
import numpy as np
import pandas as pd
import sys
sys.path.append("../..")
from mach_cad.tools.jmag import JmagDesigner
from mach_opt import InvalidDesign
from .electrical_analysis import CrossSectInnerNotchedRotor as CrossSectInnerNotchedRotor
from .electrical_analysis import CrossSectStator as CrossSectStator
from .electrical_analysis.Location2D import Location2D

EPS = 1e-2  # unit: mm


class JMagTransient2DFEA():
    def run(self,problem:'FEAProblem'):
        """
        This is where the order of operations for the 2D Transient analysis
        is defined. It calls functions implemented in the base class to 
        perform actions and then returns results.
        """
        self.init_tool()
        self.draw_components(problem.components) #Can skip this for now...
        self.create_study() #This should Switch to the study part of jmag
        self.apply_materials(problem.components) #apply the materials
        for cond in problem.conditions:
            cond.apply(self.tool)
        for sett in problem.settings:
            sett.apply(self)
        for get_res in problem.get_results:
            get_res.define(self)
        for config in problem.configs:
            config.apply(self)
        self.run_study() # run the study
        
        results = [None]*len(problem.get_results)
        results=self.collect_results(problem.get_results) # collect results
        self.clear_tool()
        return results
    
    def init_tool(self):
        self.tool=JmagDesigner()
        self.tool.open()
        
    def draw_components(self,components) :
        ####################################################
        # Adding parts object
        ####################################################
        self.rotorCore = CrossSectInnerNotchedRotor.CrossSectInnerNotchedRotor(
            name='NotchedRotor',
            mm_d_m=self.machine_variant.d_m * 1e3,
            deg_alpha_m=self.machine_variant.alpha_m,  # angular span of the pole: class type DimAngular
            deg_alpha_ms=self.machine_variant.alpha_ms,  # segment span: class type DimAngular
            mm_d_ri=self.machine_variant.d_ri * 1e3,  # inner radius of rotor: class type DimLinear
            mm_r_sh=self.machine_variant.r_sh * 1e3,  # rotor iron thickness: class type DimLinear
            mm_d_mp=self.machine_variant.d_mp * 1e3,  # inter polar iron thickness: class type DimLinear
            mm_d_ms=self.machine_variant.d_ms * 1e3,  # inter segment iron thickness: class type DimLinear
            p=self.machine_variant.p,  # Set pole-pairs to 2
            s=self.machine_variant.n_m,  # Set magnet segments/pole to 4
            location=Location2D(anchor_xy=[0, 0], deg_theta=0))

        self.shaft = CrossSectInnerNotchedRotor.CrossSectShaft(name='Shaft',
                                                               notched_rotor=self.rotorCore
                                                               )

        self.rotorMagnet = CrossSectInnerNotchedRotor.CrossSectInnerNotchedMagnet(name='RotorMagnet',
                                                                                  notched_rotor=self.rotorCore
                                                                                  )

        self.stator_core = CrossSectStator.CrossSectInnerRotorStator(name='StatorCore',
                                                                     deg_alpha_st=self.machine_variant.alpha_st,
                                                                     deg_alpha_so=self.machine_variant.alpha_so,
                                                                     mm_r_si=self.machine_variant.r_si * 1e3,
                                                                     mm_d_so=self.machine_variant.d_so * 1e3,
                                                                     mm_d_sp=self.machine_variant.d_sp * 1e3,
                                                                     mm_d_st=self.machine_variant.d_st * 1e3,
                                                                     mm_d_sy=self.machine_variant.d_sy * 1e3,
                                                                     mm_w_st=self.machine_variant.w_st * 1e3,
                                                                     mm_r_st=0,  # dummy
                                                                     mm_r_sf=0,  # dummy
                                                                     mm_r_sb=0,  # dummy
                                                                     Q=self.machine_variant.Q,
                                                                     location=Location2D(anchor_xy=[0, 0], deg_theta=0)
                                                                     )

        self.coils = CrossSectStator.CrossSectInnerRotorStatorWinding(name='Coils',
                                                                      stator_core=self.stator_core)
        ####################################################
        # Drawing parts
        ####################################################
        # Rotor Core
        list_segments = self.rotorCore.draw(toolJd)
        toolJd.bMirror = False
        toolJd.iRotateCopy = self.rotorMagnet.notched_rotor.p * 2
        try:
            region1 = toolJd.prepareSection(list_segments)
        except:
            return False

        # Shaft
        list_segments = self.shaft.draw(toolJd)
        toolJd.bMirror = False
        toolJd.iRotateCopy = 1
        region0 = toolJd.prepareSection(list_segments)

        # Rotor Magnet
        list_regions = self.rotorMagnet.draw(toolJd)
        toolJd.bMirror = False
        toolJd.iRotateCopy = self.rotorMagnet.notched_rotor.p * 2
        region2 = toolJd.prepareSection(list_regions, bRotateMerge=False)

        # Sleeve
        # sleeve = CrossSectInnerNotchedRotor.CrossSectSleeve(
        #     name='Sleeve',
        #     notched_magnet=self.rotorMagnet,
        #     d_sleeve=self.machine_variant.d_sl * 1e3  # mm
        # )
        # list_regions = sleeve.draw(toolJd)
        # toolJd.bMirror = False
        # toolJd.iRotateCopy = self.rotorMagnet.notched_rotor.p * 2
        # try:
        #     regionS = toolJd.prepareSection(list_regions)
        # except:
        #     return False

        # Stator Core
        list_regions = self.stator_core.draw(toolJd)
        toolJd.bMirror = True
        toolJd.iRotateCopy = self.stator_core.Q
        region3 = toolJd.prepareSection(list_regions)

        # Stator Winding
        list_regions = self.coils.draw(toolJd)
        toolJd.bMirror = False
        toolJd.iRotateCopy = self.coils.stator_core.Q
        region4 = toolJd.prepareSection(list_regions)

        return True
    def create_study(self):
        self.tool.study=self.tool.create_study()
    def apply_materials(self, components):
        #apply the materials
        for ind, comp in enumerate(components):
            self.tool.study.SetMaterialByName(comp.name,comp.mat.name)
            for key, value in comp.mat.values.items():
                self.tool.study.GetMaterial(comp.name).SetValue(key,value)
        # How the heck to deal with this?    
        self.tool.study.GetMaterial(u"Magnet").SetDirectionXYZ(1, 0, 0)
        self.tool.study.GetMaterial(u"Magnet").SetAxisXYZ(0, 0, 1)
        self.tool.study.GetMaterial(u"Magnet").SetOriginXYZ(0, 0, 0)
        self.tool.study.GetMaterial(u"Magnet").SetPattern(u"ParallelCircular")
        #study.GetMaterial(u"Magnet").SetValue(u"StartAngle", 0.5 * self.machine_variant.alpha_m)

    def run_study(self):
        pass# run the stufy

class MaterialTemp():
    def __init__(self,name,values):
        self.name=name
        self.values=values # this is a dictionary of key value pairs for Jmag