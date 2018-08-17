#!/Users/fanzhaf/Documents/Personal/PyCharm/PycharmProjects/excel/excel/bin/python

import sys
sys.path.append('/Users/fanzhaf/Documents/Personal/PyCharm/PycharmProjects/excel/')
import click
from lib.pangpang_helper import PangPangHelper


class Options(object):
    pass


class PangPangAction(object):
    def __init__(self, options):
        self.options = options

    def __call__(self):
        helper = PangPangHelper(self.options.department, self.options.salary, self.options.destination)
        helper.run()


@click.group(chain=True)
@click.option('-v', '--verbose', count=True)
@click.pass_context
def pangpang_sing(ctx, verbose):
    options = Options()
    ctx.obj = PangPangCLI(options)


@pangpang_sing.command('sing')
@click.option('--department', 'department', required=True, help="Full path of the department file")
@click.option('--salary', 'salary', required=True, help="Full path of the salary file")
@click.option('--destination', 'destination', required=True, help="Full path of the destination directory for "
                                                                  "generated files")
@click.pass_obj
def pangpang_options(department, salary, destination):
    pangpang_sing.options.department = department
    pangpang_sing.options.salary = salary
    pangpang_sing.options.destination = destination
    PangPangAction(pangpang_sing.options)()


class PangPangCLI(object):
    def __init__(self, options):
        self.options = options


if __name__ == "__main__":
    pangpang_sing()
