import setuptools

try:
    from pip.download import PipSession
    from pip.req import parse_requirements
except ImportError:
    raise ImportError('Install pip.')

def reqs(path):
    return [str(r.req) for r in parse_requirements(path, session=PipSession())]


INSTALL_REQUIRES = reqs('requirements.txt')
TEST_REQUIRES = reqs('test-requirements.txt')


setuptools.setup(
    name='derbystatter',
    version='0.0.2',
    author='Glen Andreas',
    author_email='derby@gandreas.com',
    install_requires=INSTALL_REQUIRES,
    tests_require=TEST_REQUIRES,
)

