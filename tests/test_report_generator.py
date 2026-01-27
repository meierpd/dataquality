"""Unit tests for the report generator module."""

import pytest
from unittest.mock import Mock, MagicMock, patch
from pathlib import Path

from orsa_analysis.reporting.report_generator import ReportGenerator


@pytest.fixture
def mock_db_manager():
    """Create a mock DatabaseManager for testing."""
    mock = Mock()
    mock.get_latest_results_for_institute = Mock(return_value=[
        {
            'institute_id': 'INST001',
            'check_name': 'sst_three_years_filled',
            'outcome_bool': 1,
            'outcome_str': 'genügend'
        }
    ])
    mock.get_institute_metadata = Mock(return_value={
        'institute_id': 'INST001',
        'version': 1,
        'file_name': 'test.xlsx'
    })
    mock.get_institut_metadata_by_finmaid = Mock(return_value={
        'FINMAID': 'INST001',
        'FinmaObjektName': 'Test Institute Ltd.',
        'Aufsichtskategorie': 'Kategorie 1',
        'MitarbeiterName': 'John Doe'
    })
    return mock


@pytest.fixture
def mock_template_manager():
    """Create a mock ExcelTemplateManager for testing."""
    mock = Mock()
    mock.create_output_workbook = Mock()
    mock.write_cell_value = Mock(return_value=True)
    mock.save_workbook = Mock()
    mock.close = Mock()
    return mock


@pytest.fixture
def report_generator(mock_db_manager, mock_template_manager, tmp_path):
    """Create a ReportGenerator instance with mocked dependencies."""
    template_path = tmp_path / "template.xlsx"
    template_path.touch()
    
    output_dir = tmp_path / "reports"
    output_dir.mkdir()
    
    with patch('orsa_analysis.reporting.report_generator.ExcelTemplateManager') as mock_tmpl_class:
        mock_tmpl_class.return_value = mock_template_manager
        generator = ReportGenerator(
            db_manager=mock_db_manager,
            template_path=template_path,
            output_dir=output_dir
        )
        generator.template_manager = mock_template_manager
        return generator


class TestReportGeneratorInstitutMetadata:
    """Test cases for institut metadata integration in ReportGenerator."""

    def test_apply_institut_metadata_success(
        self, report_generator, mock_db_manager, mock_template_manager
    ):
        """Test successful application of institut metadata."""
        # Call the method
        result = report_generator._apply_institut_metadata("INST001")
        
        # Verify database was queried
        mock_db_manager.get_institut_metadata_by_finmaid.assert_called_once_with("INST001")
        
        # Verify all four fields were written
        assert mock_template_manager.write_cell_value.call_count == 4
        
        # Verify the correct cells and values
        calls = mock_template_manager.write_cell_value.call_args_list
        assert calls[0][0] == ("Auswertung", "E2", "Test Institute Ltd.")
        assert calls[1][0] == ("Auswertung", "E3", "INST001")
        assert calls[2][0] == ("Auswertung", "E4", "Kategorie 1")
        assert calls[3][0] == ("Auswertung", "E6", "John Doe")
        
        # Verify success
        assert result is True

    def test_apply_institut_metadata_not_found(
        self, report_generator, mock_db_manager, mock_template_manager
    ):
        """Test when institut metadata is not found."""
        # Mock no metadata found
        mock_db_manager.get_institut_metadata_by_finmaid.return_value = None
        
        # Call the method
        result = report_generator._apply_institut_metadata("NONEXISTENT")
        
        # Verify database was queried
        mock_db_manager.get_institut_metadata_by_finmaid.assert_called_once_with("NONEXISTENT")
        
        # Verify no cells were written
        mock_template_manager.write_cell_value.assert_not_called()
        
        # Verify failure
        assert result is False

    def test_apply_institut_metadata_partial_data(
        self, report_generator, mock_db_manager, mock_template_manager
    ):
        """Test when some metadata fields are None."""
        # Mock partial metadata
        mock_db_manager.get_institut_metadata_by_finmaid.return_value = {
            'FINMAID': 'INST001',
            'FinmaObjektName': 'Test Institute Ltd.',
            'Aufsichtskategorie': 'Kategorie 1',
            'MitarbeiterName': None  # Missing value
        }
        
        # Call the method
        result = report_generator._apply_institut_metadata("INST001")
        
        # Verify only 3 fields were written (excluding the None value)
        assert mock_template_manager.write_cell_value.call_count == 3
        
        # Verify failure (not all fields written)
        assert result is False

    def test_apply_institut_metadata_write_failure(
        self, report_generator, mock_db_manager, mock_template_manager
    ):
        """Test when writing to cell fails."""
        # Mock write failure for one cell
        mock_template_manager.write_cell_value.side_effect = [True, False, True, True]
        
        # Call the method
        result = report_generator._apply_institut_metadata("INST001")
        
        # Verify failure (not all writes succeeded)
        assert result is False

    def test_apply_institut_metadata_exception(
        self, report_generator, mock_db_manager, mock_template_manager
    ):
        """Test error handling when an exception occurs."""
        # Mock exception during database query
        mock_db_manager.get_institut_metadata_by_finmaid.side_effect = Exception("Database error")
        
        # Call the method
        result = report_generator._apply_institut_metadata("INST001")
        
        # Verify failure
        assert result is False

    def test_generate_report_includes_institut_metadata(
        self, report_generator, mock_db_manager, mock_template_manager, tmp_path
    ):
        """Test that generate_report calls _apply_institut_metadata."""
        # Create a mock source file
        source_file = tmp_path / "source.xlsx"
        source_file.touch()
        
        # Mock the _apply_institut_metadata method
        with patch.object(report_generator, '_apply_institut_metadata') as mock_apply:
            mock_apply.return_value = True
            
            # Generate report
            result = report_generator.generate_report(
                institute_id="INST001",
                source_file_path=source_file
            )
            
            # Verify _apply_institut_metadata was called
            mock_apply.assert_called_once_with("INST001")
            
            # Verify report was generated
            assert result is not None
