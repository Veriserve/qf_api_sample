import pytest


class TestSample:
    def test_pass(self):
        assert True

    def test_fail(self):
        assert False

    @pytest.mark.skip
    def test_skip(self):
        assert True

    @pytest.fixture
    def cause_error(self):
        raise Exception

    def test_error(self, cause_error):
        assert True
