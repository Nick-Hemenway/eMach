# -*- coding: utf-8 -*-
"""
Created on Thu May  5 15:24:24 2022

@author: Martin Johnson
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Mar 23 08:53:19 2022

@author: Martin Johnson
"""
from Typing import List,Protocol,Any
import mach_cad as mc
import mach_cad as me
from abc import abstractmethod

class FEAProblem:
    def __init__(self, components:'List[mc.Components]' ,
                 conditions:'List[Conditions]',
                 settings:'List[FEASetting]',
                 get_results:'List[GetResults]',
                 config: 'Any'):
        self._components=components
        self._conditions=conditions
        self._settings=settings
        self._get_results=get_results
        self._config=config
    @property
    def components(self):
        return self._components
    @property
    def conditions(self):
        return self._components
    @property
    def settings(self):
        return self._settings
    @property
    def get_results(self):
        return self._get_results
    @property
    def config(self):
        return self._config

        
class Conditions(Protocol):
    @abstractmethod
    def __init__(self,object_refs,cond_para):
        self.object_refs
        self.cond_para
    @abstractmethod
    def apply(self,solver):
        raise NotImplementedError

class RotationCondition(Conditions):
    def __init__(self,object_refs,cond_para):
        self.object_refs=object_refs
        self.speed=cond_para[0]
        self.angle=cond_para[1]
    def apply(self,tool:'BaseRotation'):
        tool.set_rotation(self.object_refs,self.speed,self.angle)

class CurrentCondition(Conditions):
    def __init__(self,object_refs,cond_para):
        self.object_refs=object_refs
        self.Amp=cond_para[0]
        self.freq=cond_para[1]
    def apply(self,solver:'BaseCurrent'):
        solver.set_current(self.object_refs,self.amp,self.freq)

class DPNVCurrentCondition(Conditions):
    def __init__(self,object_refs,cond_para):
        self.object_refs=object_refs
        self.AmpT=cond_para[0]
        self.AmpS=cond_para[1]
        self.freq=cond_para[2]
    def apply(self,solver:'BaseCurrent'):
        solver.set_DPNVcurrent(self.object_refs,self.AmpT,self.AmpS,self.freq)
        
class FEASetting(Protocol):
    @abstractmethod
    def __init__(self,setting_para):
        self.setting_para
        
    @abstractmethod
    def apply(self,solver):
        raise NotImplementedError   

class DPNVCurrentSetting(FEASetting):
    def __init__(self,object_refs,setting_para):
        self.AmpT=setting_para[0]
        self.AmpS=setting_para[1]
        self.freq=setting_para[2]
    def apply(self,solver:'BaseCurrent'):
        solver.set_DPNVcurrent(self.object_refs,self.AmpT,self.AmpS,self.freq)
        
class GetResults(Protocol):
    @abstractmethod
    def __init__(self,object_refs):
        self.object_refs
    @abstractmethod
    def define(self,solver):
        raise NotImplementedError 
    @abstractmethod
    def extract(self,solver):
        raise NotImplementedError

class GetTorque(GetResults):
    def __init__(self,object_refs):
        self.object_refs=object_refs
    def define(self,solver):
        solver.define_torque(self.object_refs) 
    def extract(self,solver):
        torque=solver.extract_torque(self.object_refs)
        return torque

class FEAAnalyzer(me.Analyzer):
    def __init__(self,solver:'SolverBase'):
        self.solver=solver
    def analyze(self,problem:'FEAProblem'):
        results=self.solver.run(problem)
        return results

class Solver(Protocol):
    @abstractmethod
    def run(self,problem):
        raise NotImplementedError
        
class JMAGBase:
    def __init__(self):
        """ 
        This is the base implementation of a tool class. This class fufills
        the contracts in mach_cad tool classes, as well as the methods needed
        by the conditions classes
        """
        pass

class JMagTransient2DFEA(JMAGBase,Solver):
    def run(self,problem:'FEAProblem'):
        """
        This is where the order of operations for the 2D Transient analysis
        is defined. It calls functions implemented in the base class to 
        perform actions and then returns results.
        """
        
        self.init_tool()
        self.draw_components(problem.components)
        self.create_study()
        self.apply_materials(problem.components)
        self.apply_conditions(problem.conditions)
        self.apply_settings(problem.settings)
        self.define_results(problem.get_results)
        self.apply_config(problem.config)
        self.run_study()
        results=self.collect_results(problem.get_results)
        return results
    
class Static2DFEAProblem(FEAProblem):
    def __init__(self, components:'List[mc.Components]' ,
                 conditions:'List[Conditions]',
                 settings:'List[FEASetting]',
                 get_results:'List[GetResults]',
                 config: 'Any'):
        self.components = components
        self.conditions = conditions
        self.settings = settings
        self.get_results=get_results
        self.config=config
    def run(self,solver:'SolverTransient2DBase'):
        results=solver.run_static_2D_FEA(self.components,
                                  self.conditions,
                                  self.settings,
                                  self.get_results,
                                  self.config)
        return results


class SolverBase(Protocol):
    pass

class SolverTransient2DBase(SolverBase,Protocol):
    @abstractmethod
    def run_static_2D_FEA(self,components,
                          conditions,
                          settings,
                          get_results,
                          config):
        raise NotImplementedError

class SolverStatic2DBase(SolverBase,Protocol):
    @abstractmethod
    def run_static_2D_FEA(self,components,
                          conditions,
                          settings,
                          get_results,
                          config):
        raise NotImplementedError