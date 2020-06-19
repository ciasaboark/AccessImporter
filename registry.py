import logging, winreg

REGISTRY_KEY_NAME = 'SOFTWARE\\ContainerTracking'

logger = logging.getLogger('registry')
key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, REGISTRY_KEY_NAME)

class Registry:
    @staticmethod
    def write_default(name: str, value: str):
        """Write a value to the Windows registry if it does not exist
        
        Args:
            name (str): The name of the value to write.
            val (str): The value to write.
        """
        logger.debug("Checking for default value for '{0}'".format(name))
        if Registry.read_key(name, None) == None:
            logger.info("Writing default value for '{0}': '{1}'".format(name, value))
            Registry.write_key(name, value)

    @staticmethod
    def write_key(name, val):
        """Write a value into the Windows registry.

        Args:
            name (str): The name of the value to write.
            val (str): The value to write.

        """
        logger.info("Inserting value '{0}' into registry as 'HKLM\\{1}\{2}'".format(val, REGISTRY_KEY_NAME, name))
        try:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, val)
        except Exception as e:
            logger.error("Unable to write value '{0}' to value name '{1}' in key 'HKLM\\{2}'".format(val, name, REGISTRY_KEY_NAME))
            logger.exception(e)
            raise e

    @staticmethod
    def read_key(name: str, def_val):
        """Read the current value for key 'name' from the Windows registry.
        
        The default value will be used if the registry could not be opened
        or if the value name does not exist for the given key
        
        Args:
            name (str): the value name to pull from the base key
            def_val (str): a default value to use if an error occurs reading from the registry
        
        Returns:
            str: The value read from the registry or def_val
        """
        logger.info("checking for key {}".format(name))
        val = def_val
        try:
            val = winreg.QueryValueEx(key, name)[0]
        except FileNotFoundError as e:
            logger.warning("Unable to read config option '{0}' from registry, using default value of '{1}'"
                .format(name, def_val))
            Registry.write_key(name, val)
            
        return val

    @staticmethod
    def write_default_opts():
        """Write default values to the Windows registry if they do not already exist    
        """
        Registry.write_default('watch', "C:\\import")
        Registry.write_default('archive', "C:\import\\archive")
        Registry.write_default('errors', "C:\\import\\archive\\errors")
        Registry.write_default('database', "C:\\db\\database.accdb")
        Registry.write_default('log_file', "C:\\db\\logs\\watcher.log")

    @staticmethod
    def close_key():
        winreg.CloseKey(key)