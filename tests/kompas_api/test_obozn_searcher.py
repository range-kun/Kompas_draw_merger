from unittest.mock import call
from unittest.mock import create_autospec
from unittest.mock import Mock
from unittest.mock import patch

import pytest

from programm.kompas_api import KompasAPI
from programm.kompas_api import NoExecutionsError
from programm.kompas_api import NotSupportedSpecTypeError
from programm.kompas_api import OboznSearcher
from programm.kompas_api import SpecType
from programm.schemas import DrawData
from programm.schemas import DrawType as DT
from programm.schemas import FilePath
from programm.schemas import SpecSectionData


class TestOboznSearcher:
    SPEC_PATH = r"C:\User\123.spw"

    @pytest.fixture()
    def preset_data(self):
        lines_data = [
            ("ABC-00.00.00 SB", "Assembly Draw", "A1"),
            ("ABC-10.00.00", "Speca", "A2"),
            ("ABC-10.00.01", "Detail", "A4"),
        ]
        return {
            "spc_description": [1, 2, 3],
            "get_line_section": [5, 15, 20],
            "get_line_obozn_name_size": lines_data,
        }

    @pytest.fixture
    def kompas_api(self, preset_data):
        kompas_api = create_autospec(KompasAPI)
        kompas_api.const = Mock()
        kompas_api.get_line_section.side_effect = preset_data["get_line_section"]
        kompas_api.get_document_api.return_value = Mock(), Mock()
        return kompas_api

    @pytest.fixture
    def spec_path(self):
        return FilePath(self.SPEC_PATH)

    @pytest.fixture
    def get_all_lines_with_correct_type(self):
        with patch.object(
            OboznSearcher, "_get_all_lines_with_correct_type", return_value=[Mock()]
        ):
            yield

    @pytest.fixture
    def regular_oformlenie(
        self, kompas_api, preset_data, get_all_lines_with_correct_type
    ):
        kompas_api.create_spc_object.return_value = (
            SpecType.REGULAR_SPEC,
            preset_data["spc_description"],
        )

    @pytest.fixture
    def group_oformlenie(
        self, kompas_api, preset_data, get_all_lines_with_correct_type
    ):
        kompas_api.create_spc_object.return_value = (
            SpecType.GROUP_SPEC,
            preset_data["spc_description"],
        )

    @pytest.fixture
    def get_line_draws_data(self, preset_data):
        with patch.object(OboznSearcher, "_get_line_obozn_name_size") as lines_data:
            lines_data.side_effect = preset_data["get_line_obozn_name_size"]
            yield

    @pytest.fixture
    def obozn_searcher_regular(
        self, spec_path, kompas_api, regular_oformlenie, get_line_draws_data
    ):
        return OboznSearcher(spec_path, kompas_api)

    @pytest.fixture
    def obozn_searcher_group(
        self, spec_path, kompas_api, group_oformlenie, get_line_draws_data
    ):
        return OboznSearcher(spec_path, kompas_api)

    @pytest.fixture
    def mock_verify_column_not_empty(self):
        with patch.object(
            OboznSearcher, "_verify_column_not_empty"
        ) as mock_verify_column:
            mock_verify_column.return_value = True
            yield mock_verify_column

    @pytest.fixture
    def get_empty_line_obozn_name_size(self):
        with patch.object(OboznSearcher, "_get_line_obozn_name_size") as lines_data:
            lines_data.side_effect = [
                (None, None, None),
            ]

    @pytest.fixture
    def mock_group_searcher_result(self):
        with patch.object(OboznSearcher, "_get_obozn_from_group_spec") as mocked_method:
            mocked_method.return_value = [Mock(), Mock()]
            yield mocked_method

    @pytest.fixture
    def patch_get_column_numbers(self):
        with patch.object(OboznSearcher, "_get_column_numbers", return_value=[499]):
            yield

    def test_not_correct_oformlenie(self, spec_path, kompas_api):
        # given: not supported doc oformlenie
        kompas_api.create_spc_object.return_value = ("random", Mock())

        # then: expect to rise exception and close document
        with pytest.raises(NotSupportedSpecTypeError):
            OboznSearcher(spec_path, kompas_api)
            kompas_api.doc.Close.assert_called_once()

    def test_need_to_select_executions(self, obozn_searcher_group):
        # then it's asked to choose execution should return True
        assert obozn_searcher_group.need_to_select_executions() is True

    def test_correct_returned_executions(
        self, obozn_searcher_group, mock_verify_column_not_empty
    ):
        with patch.object(OboznSearcher, "_get_cell_data") as mock_cell_data:
            # given: spec with 2 columns
            mock_cell_data.side_effect = ["draw_1", "-", AttributeError]

            # when: called method to get all executions
            executions = obozn_searcher_group.get_all_spec_executions()

            # then should return dict with correct values
            assert executions == {
                "draw_1": 1,
                "Базовое исполнение": 2,
                "Все исполнения": 1000,
            }

    def test_no_executions_with_exception_raise(self, obozn_searcher_group):
        with patch.object(
            OboznSearcher, "_get_cell_data", side_effect=[AttributeError]
        ):
            # given: spec without_execution

            # when: called method to get all executions
            with pytest.raises(NoExecutionsError):
                obozn_searcher_group.get_all_spec_executions()

    def test_should_user_correct_function_by_oformlenie(
        self, obozn_searcher_regular, kompas_api
    ):
        with patch(
            "programm.kompas_api.OboznSearcher._get_obozn_from_simple_specification",
            return_value=[Mock(), Mock()],
        ):
            # when: called method to get obozn from specification
            obozn_searcher_regular.get_obozn_from_specification()

            # then: should use correct method to get spec data
            obozn_searcher_regular._get_obozn_from_simple_specification.assert_called_once()

    def test_get_executions_should_use_provided_column_number(
        self, obozn_searcher_group
    ):
        with patch.object(
            OboznSearcher, "_get_obozn_from_group_spec", return_value=[Mock(), Mock()]
        ):
            obozn_searcher_group.get_obozn_from_specification(column_numbers=[333])

            # then should get column number:
            obozn_searcher_group._get_obozn_from_group_spec.assert_called_with([333])

    def test_should_look_for_column_number_if_not_provided(
        self, obozn_searcher_group, mock_group_searcher_result, patch_get_column_numbers
    ):
        with patch.object(OboznSearcher, "_get_column_numbers", return_value=[499]):
            # when: called method to get data from group spec without column number
            obozn_searcher_group.get_obozn_from_specification()

            # then should get column number:
            obozn_searcher_group._get_obozn_from_group_spec.assert_called_with([499])

    def test_simple_spec_to_get_only_assemblies(
        self, spec_path, kompas_api, regular_oformlenie, get_line_draws_data
    ):
        # given: simple spec with only assembly draws to get
        searcher = OboznSearcher(spec_path, kompas_api, without_sub_assembles=True)

        # when: need to get list of draws from it
        spec_data = searcher.get_obozn_from_specification()

        # then: should return only one element with SpecSectionData
        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                )
            ],
            [],
        )

    def test_simple_spec_to_get_all_type_of_draws(self, obozn_searcher_regular):
        # when: need to get list of draws from  simple spec
        spec_data = obozn_searcher_regular.get_obozn_from_specification()

        # then: should return all elements
        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.DETAIL,
                    draw_names=[
                        DrawData(draw_obozn="ABC-10.00.01", draw_name="Detail")
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.SPEC_DRAW,
                    draw_names=[DrawData(draw_obozn="ABC-10.00.00", draw_name="Speca")],
                ),
            ],
            [],
        )

    def test_simple_spec_to_call_internal_methods(self, obozn_searcher_regular):
        with patch.object(OboznSearcher, "_create_spec_output") as mocked_output:
            with patch.object(
                OboznSearcher, "_parse_lines_for_detail_and_spec"
            ) as mocked_parser:
                # when: called method to get spec data
                obozn_searcher_regular.get_obozn_from_specification()

                # then: shoul user _parse_lines_for_detail_and_spec
                mocked_output.assert_called_once()
                assert mocked_parser.call_count == 2

    def test_get_line_obozn_name_size(
        self, regular_oformlenie, kompas_api, spec_path, preset_data
    ):
        with patch.object(OboznSearcher, "_get_cell_data") as mock_cell_data:
            with patch.object(
                OboznSearcher, "_parse_lines_for_detail_and_spec"
            ) as mocked_parser:
                # given: regular spec to get all lines
                mock_cell_data.side_effect = (
                    "ABC-00.00.00 SB",
                    "Assembly Draw",
                    "A1",
                    "ABC-00.00.00 XX",
                    "Speca Draw",
                    "A2",
                    "ABC-10.00.01",
                    "Detail",
                    "A4",
                )
                searcher = OboznSearcher(spec_path, kompas_api)

                # when: called method to get all spec lines
                searcher.get_obozn_from_specification()

                # then: shold get lines and transfer them to lowercase and remove spaces
                assert mock_cell_data.call_count == 9
                mocked_parser.assert_has_calls(
                    [
                        call("abc-00.00.00xx", "specadraw", "a2", 15),
                        call("abc-10.00.01", "detail", "a4", 20),
                    ]
                )

    def test_get_line_obozn_name_size_procced_exception(
        self, regular_oformlenie, kompas_api, spec_path, preset_data
    ):
        with patch.object(OboznSearcher, "_get_cell_data") as mock_cell_data:
            with patch.object(
                OboznSearcher, "_parse_lines_for_detail_and_spec"
            ) as mocked_parser:
                # given: regular spec to get all lines with one not existing cell
                mock_cell_data.side_effect = (
                    "ABC-00.00.00 SB",
                    "Assembly Draw",
                    "A1",
                    "ABC-10.00.01",
                    "Detail",
                    AttributeError,
                    "ABC-00.00.00 XX",
                    "Speca Draw",
                    "A2",
                )
                searcher = OboznSearcher(spec_path, kompas_api)

                # when: called method to get all spec lines
                searcher.get_obozn_from_specification()

                # then: shold correctly proceed than line and go to next line
                assert mock_cell_data.call_count == 9
                mocked_parser.assert_has_calls(
                    [call("abc-00.00.00xx", "specadraw", "a2", 15)]
                )

    def test_not_add_details_without_draws_to_list(
        self, obozn_searcher_regular, kompas_api
    ):
        kompas_api.get_line_section.side_effect = [5, 20, 20]
        with patch.object(OboznSearcher, "_get_line_obozn_name_size") as lines_data:
            # given: list of lines with details without draws
            lines_data.side_effect = [
                ("ABC-00.00.00 SB", "Assembly Draw", "A1"),
                ("ABC-10.00.01", "Detail", "бч"),
                ("ABC-10.00.02", "гост", "БЧ"),
            ]

            # when: need to get list of draws from it
            spec_data = obozn_searcher_regular.get_obozn_from_specification()

            # then: should filter details without draws
            assert spec_data == (
                [
                    SpecSectionData(
                        draw_type=DT.ASSEMBLY_DRAW,
                        draw_names=[
                            DrawData(
                                draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                            )
                        ],
                    ),
                ],
                [],
            )

    def test_should_check_if_detail_passed_as_spec(
        self, obozn_searcher_regular, kompas_api
    ):
        kompas_api.get_line_section.side_effect = [5, 15, 20]

        with patch.object(OboznSearcher, "_get_line_obozn_name_size") as lines_data:
            # given: list of lines with spec by mistake passed as detail draw
            lines_data.side_effect = [
                ("ABC-00.00.00 SB", "Assembly Draw", "A1"),
                ("ABC-10.00.01", "Detail", "бч"),
                ("ABC-10.00.02", "Detail", "A4"),
            ]

            # when: need to add that detail to list of errors
            spec_data = obozn_searcher_regular.get_obozn_from_specification()

        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.DETAIL,
                    draw_names=[
                        DrawData(draw_obozn="ABC-10.00.02", draw_name="Detail")
                    ],
                ),
            ],
            [
                (
                    f"\nВозможно указана деталь в качестве "
                    f"спецификации -> {self.SPEC_PATH} ||| ABC-10.00.01 \n"
                )
            ],
        )

    @pytest.mark.parametrize(
        "obozn_data,result",
        [
            ("ABC-00.00.00", False),
            ("ABC-00.00.00А", False),
            ("ABC-00.00.01", True),
            ("ABC-00.00.01А", True),
        ],
    )
    def test_is_detail_method_correct_work(
        self, obozn_searcher_regular, obozn_data, result
    ):
        # when: called a method to check if detail obozn passed should return correct value
        assert obozn_searcher_regular.is_detail(obozn_data) == result

    @pytest.mark.parametrize(
        "spec_obozn,result", [("ABC-00.00.00-01", "01"), ("ABC-00.00.00А", "-")]
    )
    def test_get_obozn_execution_if_provided(
        self,
        spec_obozn,
        result,
        spec_path,
        kompas_api,
        group_oformlenie,
        mock_group_searcher_result,
    ):
        with patch.object(
            OboznSearcher, "_get_column_number_by_execution"
        ) as mocked_method:
            # given: spec_obozn to OboznSearcher with group spec type
            mocked_method.return_value = Mock()
            obozn_searcher = OboznSearcher(spec_path, kompas_api, spec_obozn=spec_obozn)

            # when: called method to get spec search result
            obozn_searcher.get_obozn_from_specification()

            # then: should call _get_column_number_by_execution with propper values
            mocked_method.assert_called_with(result)

    @pytest.mark.parametrize(
        "spec_obozn,execution,return_result",
        [
            ("ABC-00.00.00-01", "01", ["-", "01", AttributeError]),
            ("ABC-00.00.00А", "-", ["-", "01", AttributeError]),
        ],
    )
    def test_get_column_number_by_execution(
        self,
        spec_obozn,
        execution,
        return_result,
        spec_path,
        kompas_api,
        group_oformlenie,
        mock_group_searcher_result,
        mock_verify_column_not_empty,
    ):
        with patch.object(OboznSearcher, "_get_cell_data") as mock_cell_data:
            # given: spec with 2 columns
            mock_cell_data.side_effect = return_result
            obozn_searcher = OboznSearcher(spec_path, kompas_api, spec_obozn=spec_obozn)

            # when: called method to get spec data
            obozn_searcher.get_obozn_from_specification()

            # then: should return correct column number not empty
            # +1 потому что при поиске в
            # методе _get_column_number_by_execution начинаем итерироваться от 1
            mock_group_searcher_result.assert_called_with(
                [return_result.index(execution) + 1]
            )
            mock_verify_column_not_empty.assert_called_with(
                return_result.index(execution) + 1
            )

    def test_exception_if_not_provided_correct_value_to_get_column_num(
        self,
        spec_path,
        kompas_api,
        group_oformlenie,
        mock_verify_column_not_empty,
    ):
        with patch.object(OboznSearcher, "_get_cell_data") as mock_cell_data:
            # given: spec without correct execution
            return_results = ["05", "06", AttributeError]
            mock_cell_data.side_effect = return_results
            obozn_searcher = OboznSearcher(
                spec_path, kompas_api, spec_obozn="ABC-00.00.00"
            )

            # when: called method to get spec data
            with pytest.raises(NoExecutionsError):
                obozn_searcher.get_obozn_from_specification()

    def test_group_spec_to_get_only_assemblies(
        self,
        spec_path,
        kompas_api,
        group_oformlenie,
        get_line_draws_data,
        patch_get_column_numbers,
    ):
        # given: group spec with only assembly draws to get
        searcher = OboznSearcher(spec_path, kompas_api, without_sub_assembles=True)

        # when: need to get list of draws from it
        spec_data = searcher.get_obozn_from_specification()

        # then: should return only one element with SpecSectionData
        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                )
            ],
            [],
        )

    def test_group_spec_to_get_all_type_of_draws(
        self, obozn_searcher_group, get_line_draws_data, patch_get_column_numbers
    ):
        # when: need to get list of draws from group spec
        with patch.object(
            OboznSearcher, "_get_cell_data", return_value=True
        ) as mocked_cell_data:
            with patch.object(
                OboznSearcher,
                "_get_all_lines_with_correct_type",
                return_value=[Mock()] * 3,
            ):
                spec_data = obozn_searcher_group.get_obozn_from_specification()

        # then: should return all elements
        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.DETAIL,
                    draw_names=[
                        DrawData(draw_obozn="ABC-10.00.01", draw_name="Detail")
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.SPEC_DRAW,
                    draw_names=[DrawData(draw_obozn="ABC-10.00.00", draw_name="Speca")],
                ),
            ],
            [],
        )
        assert mocked_cell_data.call_count == 3

    def test_group_spec_to_get_only_unique_rows(
        self,
        obozn_searcher_group,
        get_line_draws_data,
        patch_get_column_numbers,
    ):
        # given: list of lines with duplicate rows
        with patch.object(
            OboznSearcher, "_get_cell_data", return_value=True
        ) as mocked_cell_data:
            with patch.object(
                OboznSearcher,
                "_get_all_lines_with_correct_type",
                return_value=[Mock()] * 5,
            ):
                with patch.object(
                    OboznSearcher, "_get_line_obozn_name_size"
                ) as lines_data:
                    lines_data.side_effect = [
                        ("ABC-00.00.00 SB", "Assembly Draw", "A1"),
                        ("ABC-10.00.00", "Speca", "A2"),
                        ("ABC-10.00.00", "Speca", "A2"),
                        ("ABC-10.00.01", "Detail", "A4"),
                        ("ABC-10.00.01", "Detail", "A4"),
                    ]

                    # when: need to get list of draws from  group spec
                    spec_data = obozn_searcher_group.get_obozn_from_specification()

        # then: should return only unique elements
        assert spec_data == (
            [
                SpecSectionData(
                    draw_type=DT.ASSEMBLY_DRAW,
                    draw_names=[
                        DrawData(
                            draw_obozn="ABC-00.00.00 SB", draw_name="Assembly Draw"
                        )
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.DETAIL,
                    draw_names=[
                        DrawData(draw_obozn="ABC-10.00.01", draw_name="Detail")
                    ],
                ),
                SpecSectionData(
                    draw_type=DT.SPEC_DRAW,
                    draw_names=[DrawData(draw_obozn="ABC-10.00.00", draw_name="Speca")],
                ),
            ],
            [],
        )
        assert mocked_cell_data.call_count == 3

    def test_group_searcher_call_internal_methods(
        self, obozn_searcher_group, patch_get_column_numbers
    ):
        with patch.object(
            OboznSearcher, "_parse_lines_for_detail_and_spec", return_value=True
        ) as mocked_parser:
            with patch.object(OboznSearcher, "_create_spec_output") as mocked_output:
                # when: called method to get spec data
                obozn_searcher_group.get_obozn_from_specification()

                # then: shoul user _parse_lines_for_detail_and_spec
                mocked_parser.assert_called_once()
                mocked_output.assert_called_once()
