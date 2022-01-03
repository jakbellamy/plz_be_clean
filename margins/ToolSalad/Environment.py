class Environment(object):
    def __init__(self, env_path='.env'):
        with open(env_path) as file:
            var_list = list(filter(bool, file.read().split('\n')))

        self._vars = {x.split('=')[0].split(' ')[-1]: x.split('"')[-2] for x in var_list}
        self._available = self._vars.keys()

        for (key, value) in self._vars.items():
            self.__dict__[key] = value
