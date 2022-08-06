#!/usr/bin/env python
""" Handle parameters for the CSH High Holiday Honors programs.
    Parameters can come on the command line or from the tmstats.yml file.

    Usage:
       Create the cshparse instance.
       If additional parameters are needed, add them to cshparse.parser.
       Call tmparser.parse() to handle common parameters.
       Interpret other parameters in self.args as needed.
    """
from __future__ import print_function
import argparse, yaml, os

class CustomFormatter(argparse.ArgumentDefaultsHelpFormatter, argparse.RawDescriptionHelpFormatter):
    pass


class Singleton(object):
    def __new__(type, *args, **kwargs):
        if not '_the_instance' in type.__dict__:
            type._the_instance = object.__new__(type)
        return type._the_instance

class cshparse(Singleton):
    def __init__(self, description='Part of the CSH High Holiday Honors process', YMLfile='honors.yml', includedbparms=True, customformatter=True, **kwargs):
        if self.__dict__.get('parms', False):
            return
        if customformatter is True:
            formatter_class = CustomFormatter
        elif customformatter:
            formatter_class = customformatter
        else:
            formatter_class = None
        if formatter_class:
            self.parser = argparse.ArgumentParser(description=description, formatter_class=formatter_class, **kwargs)
        else:
            self.parser = argparse.ArgumentParser(description=description)
        self.parser.add_argument('--YMLfile',  help="YML file with information for this program", default=YMLfile)
        if includedbparms:
            self.parser.add_argument('--dbname', help="MySQL database to use", dest='dbname')
            self.parser.add_argument('--dbhost', help="host for MySQL database", dest='dbhost')
            self.parser.add_argument('--dbuser', help="user for MySQL database", dest='dbuser')
            self.parser.add_argument('--dbpass', help="password for MySQL database", dest='dbpass')

    def __repr__(self):
        return '\n'.join(['%s: "%s"' % (k, self.__dict__[k]) for k in self.__dict__ if k != 'parser'])

    def add_argument(self, *args, **kwargs):
        self.parser.add_argument(*args, **kwargs)

    def add_argument_group(self, *args, **kwargs):
        return self.parser.add_argument_group(*args, **kwargs)

    def add_mutually_exclusive_group(self, *args, **kwargs):
        return self.parser.add_mutually_exclusive_group(*args, **kwargs)

    def parse(self):

        # Parameters are put directly into this object, based on their name in the YML file
        #   or the command line.
        # NOTE:  Parameters with default values which evaluate to TRUE will ALWAYS override the file!
        #
        # self.ymlvalues is the result of reading the YML file
        # self.args is the result from the parser

        # Parse the command line (in case the YMLfile has been overridden)
        self.args = self.parser.parse_args()

        # Set values from the YML file
        self.ymlvalues= yaml.load(open(self.args.YMLfile,'r'), Loader=yaml.FullLoader)
        for name in self.ymlvalues:
            self.__dict__[name] = self.ymlvalues[name]


        # Override with non-false values from the command line (or the default).
        # If no value is in the YMLfile, use the command line or default whether it's true or false.
        args = vars(self.args)
        for name in args.keys():
            if args[name] or name not in self.__dict__:
                self.__dict__[name] = args[name]

        # And handle dbhost specially to make sure it exists:
        if 'dbhost' not in self.__dict__ or not self.dbhost:
            self.dbhost = 'localhost'


