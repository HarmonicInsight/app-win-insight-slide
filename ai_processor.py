"""
AI Processor - 複数のAI APIに対応したテキスト処理
"""
import json
import os
from abc import ABC, abstractmethod
from typing import Optional


class AIProvider(ABC):
    """AI プロバイダーの基底クラス"""
    
    @abstractmethod
    def process(self, prompt: str, text: str) -> str:
        """テキストを処理して結果を返す"""
        pass
    
    @abstractmethod
    def process_json(self, prompt: str, json_data: dict) -> dict:
        """JSON全体を処理して結果を返す"""
        pass


class OpenAIProvider(AIProvider):
    """OpenAI API"""
    
    def __init__(self, api_key: str, model: str = "gpt-4o"):
        self.api_key = api_key
        self.model = model
    
    def process(self, prompt: str, text: str) -> str:
        try:
            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": text}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"[エラー: {str(e)}]"
    
    def process_json(self, prompt: str, json_data: dict) -> dict:
        try:
            import openai
            client = openai.OpenAI(api_key=self.api_key)
            
            full_prompt = f"""{prompt}

以下のJSONのtextフィールドを処理してください。
JSON構造は変更せず、textフィールドのみ更新して返してください。
必ず有効なJSONで返答してください。"""
            
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": full_prompt},
                    {"role": "user", "content": json.dumps(json_data, ensure_ascii=False)}
                ],
                response_format={"type": "json_object"}
            )
            
            result_text = response.choices[0].message.content
            return json.loads(result_text)
        except Exception as e:
            print(f"AI処理エラー: {e}")
            return json_data


class ClaudeProvider(AIProvider):
    """Anthropic Claude API"""
    
    def __init__(self, api_key: str, model: str = "claude-sonnet-4-20250514"):
        self.api_key = api_key
        self.model = model
    
    def process(self, prompt: str, text: str) -> str:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=self.api_key)
            
            response = client.messages.create(
                model=self.model,
                max_tokens=4096,
                messages=[
                    {"role": "user", "content": f"{prompt}\n\n{text}"}
                ]
            )
            return response.content[0].text
        except Exception as e:
            return f"[エラー: {str(e)}]"
    
    def process_json(self, prompt: str, json_data: dict) -> dict:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=self.api_key)
            
            full_prompt = f"""{prompt}

以下のJSONのtextフィールドを処理してください。
JSON構造は変更せず、textフィールドのみ更新して返してください。
必ず有効なJSONのみで返答してください。コードブロックや説明は不要です。

{json.dumps(json_data, ensure_ascii=False)}"""
            
            response = client.messages.create(
                model=self.model,
                max_tokens=4096,
                messages=[
                    {"role": "user", "content": full_prompt}
                ]
            )
            
            result_text = response.content[0].text
            # コードブロックが含まれている場合は除去
            if "```json" in result_text:
                result_text = result_text.split("```json")[1].split("```")[0]
            elif "```" in result_text:
                result_text = result_text.split("```")[1].split("```")[0]
            
            return json.loads(result_text.strip())
        except Exception as e:
            print(f"AI処理エラー: {e}")
            return json_data


class MockProvider(AIProvider):
    """テスト用モックプロバイダー"""
    
    def process(self, prompt: str, text: str) -> str:
        return f"[処理済] {text}"
    
    def process_json(self, prompt: str, json_data: dict) -> dict:
        # textフィールドに [処理済] を付加
        result = json.loads(json.dumps(json_data))  # deep copy
        for slide in result.get("slides", []):
            for shape in slide.get("shapes", []):
                shape["text"] = f"[処理済] {shape['text']}"
        return result


class AIProcessor:
    """AI処理のメインクラス"""
    
    PROVIDERS = {
        "openai": OpenAIProvider,
        "claude": ClaudeProvider,
        "mock": MockProvider
    }
    
    def __init__(self):
        self.provider: Optional[AIProvider] = None
        self.presets: dict = {
            "翻訳（英語）": "以下のテキストを英語に翻訳してください。",
            "翻訳（日本語）": "以下のテキストを日本語に翻訳してください。",
            "要約": "以下のテキストを簡潔に要約してください。",
            "敬語変換": "以下のテキストを丁寧な敬語に変換してください。",
            "カジュアル変換": "以下のテキストをカジュアルな表現に変換してください。",
            "校正": "以下のテキストの誤字脱字を修正してください。",
            "建設用語統一": "以下のテキストを建設業界の専門用語を使用した表現に統一してください。",
        }
    
    def set_provider(self, provider_name: str, api_key: str = "", **kwargs):
        """プロバイダーを設定"""
        if provider_name == "mock":
            self.provider = MockProvider()
        elif provider_name in self.PROVIDERS:
            self.provider = self.PROVIDERS[provider_name](api_key, **kwargs)
        else:
            raise ValueError(f"Unknown provider: {provider_name}")
    
    def process_text(self, prompt: str, text: str) -> str:
        """単一テキストを処理"""
        if self.provider is None:
            raise RuntimeError("プロバイダーが設定されていません")
        return self.provider.process(prompt, text)
    
    def process_json(self, prompt: str, json_data: dict) -> dict:
        """JSON全体を処理"""
        if self.provider is None:
            raise RuntimeError("プロバイダーが設定されていません")
        return self.provider.process_json(prompt, json_data)
    
    def get_presets(self) -> dict:
        """プリセット一覧を取得"""
        return self.presets
    
    def add_preset(self, name: str, prompt: str):
        """プリセットを追加"""
        self.presets[name] = prompt


# テスト用
if __name__ == "__main__":
    processor = AIProcessor()
    processor.set_provider("mock")
    
    test_data = {
        "slides": [
            {
                "slide": 1,
                "shapes": [
                    {"id": 1, "name": "Title", "text": "会社概要"}
                ]
            }
        ]
    }
    
    result = processor.process_json("テスト", test_data)
    print(json.dumps(result, ensure_ascii=False, indent=2))
