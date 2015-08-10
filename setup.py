
from distutils.core import setup
from flyingpandas import __version__

setup(
    name='flyingpandas',
    version=__version__,
    author='Robert West',
    author_email='robert.david.west@gmail.com',
    packages=['flyingpandas'],
    scripts=[],
    description='''`flyingpandas.to_excel` is just like 
                   `pandas.DataFrame.to_excel`. Except that you can make things 
                   really pretty really quickly. `flyingpandas` uses to 
                   `xlwings` to add formatting to your excel files after you've 
                   used `pandas.to_excel`''',
    long_description=open('README.md').read(),
    install_requires=['pandas', 'xlwings'],    
)