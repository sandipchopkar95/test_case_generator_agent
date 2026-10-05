from types import SimpleNamespace

import pytest

import generate_tests


class FakeAnthropicMessages:
    def __init__(self, response):
        self.response = response
        self.request = None

    def create(self, **kwargs):
        self.request = kwargs
        return self.response


class FakeOpenAI:
    def __init__(self, **kwargs):
        self.kwargs = kwargs


def test_direct_provider_configuration_uses_provider_specific_model_defaults():
    assert generate_tests.DEFAULT_MODELS["openai"] == "gpt-4o"
    assert generate_tests.DEFAULT_MODELS["anthropic"] == "claude-sonnet-4-6"
    assert generate_tests.PROVIDER_API_KEY_ENV["openai"] == "OPENAI_API_KEY"
    assert generate_tests.PROVIDER_API_KEY_ENV["anthropic"] == "ANTHROPIC_API_KEY"


def test_create_model_client_requires_the_selected_provider_key(monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)

    with pytest.raises(ValueError, match="OPENAI_API_KEY"):
        generate_tests.create_model_client("openai")


def test_create_openai_client_uses_direct_api(monkeypatch):
    monkeypatch.setattr(generate_tests, "OpenAI", FakeOpenAI)

    client = generate_tests.create_model_client("openai", "openai-test-key", timeout=17)

    assert client.provider == "openai"
    assert client.sdk_client.kwargs == {
        "api_key": "openai-test-key",
        "timeout": 17,
        "max_retries": 0,
    }


def test_anthropic_adapter_translates_tools_images_and_response():
    tool_response = SimpleNamespace(
        content=[
            SimpleNamespace(
                type="tool_use",
                id="tool-1",
                name="submit_test_cases",
                input={"test_cases": []},
            ),
            SimpleNamespace(type="text", text="done"),
        ]
    )
    messages = FakeAnthropicMessages(tool_response)
    sdk_client = SimpleNamespace(messages=messages)
    image = generate_tests._multimodal_user_content("Review this image", [("screen.png", b"image-bytes")])

    response = generate_tests._anthropic_completion(
        sdk_client,
        model="claude-sonnet-4-6",
        max_tokens=1200,
        messages=[
            {"role": "system", "content": "Follow the schema."},
            {"role": "user", "content": image},
        ],
        tools=[generate_tests.TEST_CASE_TOOL],
        tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
        response_format={"type": "json_object"},
    )

    assert messages.request["system"] == "Follow the schema."
    assert messages.request["messages"][0]["content"][0] == {
        "type": "text",
        "text": "Review this image",
    }
    assert messages.request["messages"][0]["content"][1]["source"]["media_type"] == "image/png"
    assert messages.request["tools"][0]["name"] == "submit_test_cases"
    assert messages.request["tool_choice"] == {"type": "tool", "name": "submit_test_cases"}
    assert response.choices[0].message.content == "done"
    assert response.choices[0].message.tool_calls[0].function.name == "submit_test_cases"
    assert generate_tests._extract_tool_result(response) == {"test_cases": []}
