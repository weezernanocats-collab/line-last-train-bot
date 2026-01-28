"""
マルチエージェントリーダーシステム

1人のリーダー（武将口調）と2人の部下（足軽口調）による協調的問題解決システム。
リーダーが方向性を決定し、部下がそれぞれの性格に基づいたアイデアを提供。
最終的にリーダーが多目的最適化の観点から結論を導出する。
"""

import os
from typing import Optional
from dataclasses import dataclass
from anthropic import Anthropic


@dataclass
class AgentResponse:
    """エージェントからの応答を格納するデータクラス"""
    agent_name: str
    role: str
    content: str


class BaseAgent:
    """エージェントの基底クラス"""

    def __init__(self, name: str, role: str, system_prompt: str, model: str = "claude-sonnet-4-20250514"):
        self.name = name
        self.role = role
        self.system_prompt = system_prompt
        self.model = model
        self.client = Anthropic()

    def think(self, prompt: str, context: Optional[str] = None) -> AgentResponse:
        """与えられたプロンプトに対して思考し、応答を返す"""
        messages = []

        if context:
            messages.append({
                "role": "user",
                "content": f"【背景情報】\n{context}\n\n【課題】\n{prompt}"
            })
        else:
            messages.append({
                "role": "user",
                "content": prompt
            })

        response = self.client.messages.create(
            model=self.model,
            max_tokens=2048,
            system=self.system_prompt,
            messages=messages
        )

        return AgentResponse(
            agent_name=self.name,
            role=self.role,
            content=response.content[0].text
        )


class LeaderAgent(BaseAgent):
    """リーダーエージェント（武将口調）"""

    SYSTEM_PROMPT = """汝は戦国時代の武将「信玄」である。知略と武勇を兼ね備え、配下の者たちを率いて戦いに臨む大将じゃ。

【口調の特徴】
- 「〜じゃ」「〜であるぞ」「〜せよ」など武将らしい言葉遣い
- 「うむ」「ふむ」などの相槌
- 配下には「者ども」「足軽ども」と呼びかける
- 威厳と慈愛を持って部下に接する

【行動指針】
- まず問題の本質を見極める
- 部下に明確な指示を与える
- 部下の意見を尊重しつつも、最終決断は自らが下す
- 多目的最適化（複数の観点のバランス）を重視する

【多目的最適化の観点】
結論を出す際は以下の観点を考慮せよ：
1. 実現可能性（現実的に実行できるか）
2. 効果の大きさ（得られる成果）
3. リスク（失敗した場合の損害）
4. 革新性（新しい価値を生み出すか）
5. 持続可能性（長期的に維持できるか）

回答は簡潔かつ明瞭に。戦略的思考を持って臨め。"""

    def __init__(self):
        super().__init__(
            name="信玄",
            role="大将",
            system_prompt=self.SYSTEM_PROMPT
        )

    def analyze_problem(self, problem: str) -> AgentResponse:
        """問題を分析し、方向性を定める"""
        prompt = f"""以下の課題について、方向性を定めよ。

【課題】
{problem}

以下の形式で回答せよ：
1. 問題の本質
2. 検討すべき観点
3. 足軽どもへの指示"""

        return self.think(prompt)

    def give_orders(self, problem: str, analysis: str) -> tuple[str, str]:
        """部下への指示を生成"""
        prompt = f"""先の分析を踏まえ、二人の足軽に具体的な指示を出せ。

【元の課題】
{problem}

【先の分析】
{analysis}

以下の形式で、二人の足軽それぞれに異なる観点から検討させる指示を出せ：

【太郎への指示】
（慎重な性格の太郎には、リスクや安全性の観点から検討させよ）

【次郎への指示】
（大胆な性格の次郎には、革新性や効果の観点から検討させよ）"""

        response = self.think(prompt)
        content = response.content

        # 指示を分割（簡易的な解析）
        taro_order = ""
        jiro_order = ""

        if "【太郎への指示】" in content and "【次郎への指示】" in content:
            parts = content.split("【次郎への指示】")
            taro_part = parts[0]
            jiro_order = parts[1].strip() if len(parts) > 1 else ""

            if "【太郎への指示】" in taro_part:
                taro_order = taro_part.split("【太郎への指示】")[1].strip()

        return taro_order, jiro_order

    def synthesize_conclusion(self, problem: str, taro_response: str, jiro_response: str) -> AgentResponse:
        """部下の意見を統合し、多目的最適化の観点から結論を導出"""
        prompt = f"""部下どもの報告を聞き、最終的な結論を導け。

【元の課題】
{problem}

【太郎（慎重派）の報告】
{taro_response}

【次郎（大胆派）の報告】
{jiro_response}

以下の多目的最適化の観点から、バランスの取れた最終結論を出せ：
1. 実現可能性
2. 効果の大きさ
3. リスク
4. 革新性
5. 持続可能性

【回答形式】
まず部下どもの意見を評価し、その後最終結論を述べよ。
良い点は取り入れ、問題点は指摘し、総合的な判断を下せ。"""

        return self.think(prompt)


class SubordinateAgent(BaseAgent):
    """部下エージェント（足軽口調）の基底クラス"""
    pass


class TaroAgent(SubordinateAgent):
    """太郎エージェント（慎重な足軽）"""

    SYSTEM_PROMPT = """おいらは足軽の「太郎」でござる。お殿様に仕える忠実な兵士じゃ。

【口調の特徴】
- 「〜でござる」「〜でありますな」など足軽らしい丁寧な言葉遣い
- 「へい」「はっ」などの返事
- お殿様には敬意を持って接する
- 控えめだが芯のある態度

【性格】
- 慎重で石橋を叩いて渡るタイプ
- リスクを重視し、最悪の事態を想定する
- 堅実で着実な方法を好む
- 時に慎重すぎて機会を逃すこともある
- 細部に気を配り、見落としがちな問題点を指摘する

【思考パターン】
- まず「何がうまくいかない可能性があるか」を考える
- 安全マージンを十分に取った提案をする
- 過去の失敗例から学ぶことを重視
- 実現可能性と持続可能性を重視

回答は謙虚に、しかし自分の意見はしっかり述べよ。"""

    def __init__(self):
        super().__init__(
            name="太郎",
            role="足軽（慎重派）",
            system_prompt=self.SYSTEM_PROMPT
        )

    def execute_order(self, order: str, original_problem: str) -> AgentResponse:
        """お殿様の命令を実行"""
        prompt = f"""お殿様より命令が下された。精一杯お答えするでござる。

【お殿様からの命令】
{order}

【元々の課題】
{original_problem}

慎重な観点から、以下について報告せよ：
1. 考えられるリスクや問題点
2. それを回避・軽減する方法
3. 堅実な解決策の提案

正直に、思うところを述べよ。"""

        return self.think(prompt)


class JiroAgent(SubordinateAgent):
    """次郎エージェント（大胆な足軽）"""

    SYSTEM_PROMPT = """おいらは足軽の「次郎」でござる。お殿様に仕える勇敢な兵士じゃ！

【口調の特徴】
- 「〜でござるよ！」「〜じゃないですかい！」など元気な言葉遣い
- 「おう！」「よっしゃ！」などの威勢のいい返事
- お殿様には敬意を持ちつつも、少し砕けた態度
- 明るく前向きな姿勢

【性格】
- 大胆で挑戦を恐れないタイプ
- チャンスを重視し、攻めの姿勢を取る
- 革新的で斬新な方法を好む
- 時に大胆すぎて失敗することもある
- 新しいアイデアを次々と思いつく

【思考パターン】
- まず「どうすれば大きな成果が得られるか」を考える
- 既存の枠にとらわれない発想をする
- 未来の可能性を重視
- 効果の大きさと革新性を重視

回答は元気よく、自信を持って意見を述べよ！"""

    def __init__(self):
        super().__init__(
            name="次郎",
            role="足軽（大胆派）",
            system_prompt=self.SYSTEM_PROMPT
        )

    def execute_order(self, order: str, original_problem: str) -> AgentResponse:
        """お殿様の命令を実行"""
        prompt = f"""お殿様より命令が下されたでござるよ！全力でお答えするでござる！

【お殿様からの命令】
{order}

【元々の課題】
{original_problem}

大胆な観点から、以下について報告せよ：
1. 大きな成果を得るためのアイデア
2. 革新的・斬新なアプローチ
3. チャンスを最大化する提案

思い切って、アイデアをぶつけてみせるでござる！"""

        return self.think(prompt)


class MultiAgentLeaderSystem:
    """マルチエージェントリーダーシステム"""

    def __init__(self):
        self.leader = LeaderAgent()
        self.taro = TaroAgent()
        self.jiro = JiroAgent()

    def solve(self, problem: str, verbose: bool = True) -> dict:
        """
        問題を解決する

        Args:
            problem: 解決すべき問題・課題
            verbose: 途中経過を表示するかどうか

        Returns:
            各エージェントの応答と最終結論を含む辞書
        """
        results = {
            "problem": problem,
            "leader_analysis": None,
            "taro_response": None,
            "jiro_response": None,
            "final_conclusion": None
        }

        # ステップ1: リーダーが問題を分析
        if verbose:
            print("=" * 60)
            print("【第一段階】大将による問題分析")
            print("=" * 60)

        analysis = self.leader.analyze_problem(problem)
        results["leader_analysis"] = analysis

        if verbose:
            print(f"\n🏯 {analysis.agent_name}（{analysis.role}）：")
            print(analysis.content)
            print()

        # ステップ2: リーダーが部下に指示を出す
        if verbose:
            print("=" * 60)
            print("【第二段階】足軽への指示")
            print("=" * 60)

        taro_order, jiro_order = self.leader.give_orders(problem, analysis.content)

        if verbose:
            print(f"\n📜 太郎への指示：\n{taro_order}")
            print(f"\n📜 次郎への指示：\n{jiro_order}")
            print()

        # ステップ3: 部下がそれぞれ検討
        if verbose:
            print("=" * 60)
            print("【第三段階】足軽どもの報告")
            print("=" * 60)

        taro_response = self.taro.execute_order(taro_order, problem)
        results["taro_response"] = taro_response

        if verbose:
            print(f"\n⚔️ {taro_response.agent_name}（{taro_response.role}）：")
            print(taro_response.content)
            print()

        jiro_response = self.jiro.execute_order(jiro_order, problem)
        results["jiro_response"] = jiro_response

        if verbose:
            print(f"\n⚔️ {jiro_response.agent_name}（{jiro_response.role}）：")
            print(jiro_response.content)
            print()

        # ステップ4: リーダーが結論を出す
        if verbose:
            print("=" * 60)
            print("【最終段階】大将による結論")
            print("=" * 60)

        conclusion = self.leader.synthesize_conclusion(
            problem,
            taro_response.content,
            jiro_response.content
        )
        results["final_conclusion"] = conclusion

        if verbose:
            print(f"\n🏯 {conclusion.agent_name}（{conclusion.role}）の最終結論：")
            print(conclusion.content)
            print()

        return results

    def get_summary(self, results: dict) -> str:
        """結果の要約を取得"""
        summary = f"""
╔══════════════════════════════════════════════════════════════╗
║          マルチエージェントリーダーシステム - 結果要約         ║
╚══════════════════════════════════════════════════════════════╝

【課題】
{results['problem']}

【大将（信玄）の分析】
{results['leader_analysis'].content if results['leader_analysis'] else 'なし'}

【太郎（慎重派）の提案】
{results['taro_response'].content if results['taro_response'] else 'なし'}

【次郎（大胆派）の提案】
{results['jiro_response'].content if results['jiro_response'] else 'なし'}

【最終結論】
{results['final_conclusion'].content if results['final_conclusion'] else 'なし'}
"""
        return summary


# メイン実行用
if __name__ == "__main__":
    import sys

    # コマンドライン引数からプロンプトを取得、なければデフォルト
    if len(sys.argv) > 1:
        problem = " ".join(sys.argv[1:])
    else:
        problem = "新しいWebサービスを立ち上げたいが、技術スタックの選定で迷っている。スピード重視か品質重視か、どちらの方針で進めるべきか。"

    print("\n" + "🏰" * 30)
    print("  マルチエージェントリーダーシステム起動")
    print("🏰" * 30 + "\n")

    system = MultiAgentLeaderSystem()
    results = system.solve(problem)

    print("\n" + "=" * 60)
    print("処理完了")
    print("=" * 60)
