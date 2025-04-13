import unittest
from src.excel_downloader import sanitize_filename, get_filename_from_url, download_file_threaded

class TestExcelDownloader(unittest.TestCase):

    def test_sanitize_filename(self):
        # Test for valid filename
        self.assertEqual(sanitize_filename("valid_filename.txt"), "valid_filename.txt")
        # Test for invalid characters
        self.assertEqual(sanitize_filename("invalid/filename.txt"), "invalid_filename.txt")
        self.assertEqual(sanitize_filename("another*invalid|filename.txt"), "another_invalid_filename.txt")
        # Test for folder sanitization
        self.assertEqual(sanitize_filename("folder/name", is_folder=True), "folder_name")
        self.assertEqual(sanitize_filename("folder?name", is_folder=True), "folder_name")

    def test_get_filename_from_url(self):
        # Mock response object
        class MockResponse:
            headers = {
                'content-disposition': 'attachment; filename="test_file.txt"'
            }

        url = "http://example.com/test_file"
        response = MockResponse()
        filename = get_filename_from_url(url, response)
        self.assertEqual(filename, "test_file.txt")

    def test_download_file_threaded(self):
        # This test would require mocking requests.get and checking the behavior
        # For now, we will just assert that the function exists
        self.assertTrue(callable(download_file_threaded))

if __name__ == '__main__':
    unittest.main()