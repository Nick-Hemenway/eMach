import sys
import pickle
import pygmo as pg

sys.path.append("..")

import des_opt as do


class DataHandler(do.DataHandler):
    def __init__(self, archive_filepath, designer_filepath):
        self.archive_filepath = archive_filepath
        self.designer_filepath = designer_filepath

    def save_to_archive(self, x, design, full_results, objs):
        # assign relevant data to OptiData class attributes
        opti_data = do.OptiData(x=x, design=design, full_results=full_results, objs=objs)
        # write to pkl file. 'ab' indicates binary append
        with open(self.archive_filepath, 'ab') as archive:
            pickle.dump(opti_data, archive, -1)

    def load_from_archive(self):
        with open(self.archive_filepath, 'rb') as f:
            while 1:
                try:
                    yield pickle.load(f)  # use generator
                except EOFError:
                    break

    def save_designer(self, designer):
        with open(self.designer_filepath, 'wb') as des:
            pickle.dump(designer, des, -1)
            
    def get_archive_data(self):
        archive = self.load_from_archive()
        fitness = []
        free_vars = []
        for data in archive:
            fitness.append(data.objs)
            free_vars.append(data.x)
        return fitness, free_vars
    
    def get_pareto_designs(self):
        archive = self.load_from_archive()
        fitness, free_vars = self.get_archive_data()
        
        ndf, dl, dc, ndr = pg.fast_non_dominated_sorting(fitness)
        fronts_index = ndf[0]
        
        i = 0
        for data in archive:
            if i in fronts_index:
                yield data
            i = i+1
    
    def get_pareto_data(self):
        archive = self.get_pareto_designs()
        fitness = []
        free_vars = []
        rated_power = []
        for data in archive:
            fitness.append(data.objs)
            free_vars.append(data.x)
            
            final_state = data.full_results[-1][-1]
            machine = final_state.design.machine
            em_results = final_state.conditions.em
            power = em_results['torque_avg'] * machine.mech_omega /1000
            rated_power.append(power)
            
        return fitness, free_vars, rated_power

    def get_designs(self):
        archive = self.load_from_archive()
        i = 0
        for data in archive:
            final_state = data.full_results[-1][-1]
            Ea = final_state.conditions.em["Ea"]
            i = i+1
            if data.objs[0]<-6.5 and Ea<5 and data.objs[2]<5:
                print(final_state.design.machine.name)
                print(data.objs, Ea)
                
    