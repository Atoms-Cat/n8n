from outlook_client.errors import GraphAPIError

def test_graph_api_error_formatting():
    err = GraphAPIError(404, "Not Found")
    assert err.status_code == 404
    assert str(err) == "404: Not Found"
