from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class CopyRule:
    source_col: str
    target_col: str


@dataclass(frozen=True)
class ConditionalCopyRule:
    condition_source_col: str
    allowed_values: tuple[str, ...]
    source_col: str
    target_col: str


@dataclass(frozen=True)
class AutoNumberRule:
    target_col: str
    start_number: int = 1


@dataclass(frozen=True)
class FixedValueRule:
    target_col: str
    value: Any
    only_if_blank: bool = False


@dataclass(frozen=True)
class ValueMapCopyRule:
    source_col: str
    target_col: str
    value_map: dict[Any, Any]
    default_value: Any = None


@dataclass(frozen=True)
class MappingConfig:
    always_copy_rules: tuple[CopyRule, ...] = field(default_factory=tuple)
    conditional_copy_rules: tuple[ConditionalCopyRule, ...] = field(default_factory=tuple)
    auto_number_rules: tuple[AutoNumberRule, ...] = field(
        default_factory=lambda: (AutoNumberRule(target_col="B"),)
    )
    fixed_value_rules: tuple[FixedValueRule, ...] = field(default_factory=tuple)
    value_map_copy_rules: tuple[ValueMapCopyRule, ...] = field(default_factory=tuple)
    formula_autofill_enabled: bool = True

    @property
    def is_configured(self) -> bool:
        return bool(
            self.always_copy_rules
            or self.conditional_copy_rules
            or self.fixed_value_rules
            or self.value_map_copy_rules
        )

    def required_monthly_columns(self) -> set[str]:
        columns: set[str] = set()

        for rule in self.always_copy_rules:
            columns.add(rule.source_col)

        for rule in self.conditional_copy_rules:
            columns.add(rule.condition_source_col)
            columns.add(rule.source_col)

        for rule in self.value_map_copy_rules:
            columns.add(rule.source_col)

        return columns

    def required_master_columns(self) -> set[str]:
        columns: set[str] = set()

        for rule in self.always_copy_rules:
            columns.add(rule.target_col)

        for rule in self.conditional_copy_rules:
            columns.add(rule.target_col)

        for rule in self.auto_number_rules:
            columns.add(rule.target_col)

        for rule in self.fixed_value_rules:
            columns.add(rule.target_col)

        for rule in self.value_map_copy_rules:
            columns.add(rule.target_col)

        return columns


# 値が入らなかったときだけ "-" を入れる列
HYPHEN_IF_BLANK_TARGETS = (
    "Y",
    "Z",
    "AA",
    "AB",
    "AC",
    "AD",
    "AE",
    "AF",
    "AG",
    "AH",
    "AI",
    "AJ",
    "AK",
    "AT",
    "AZ",
    "BA",
    "BB",
    "BI",
    "BN",
    "BO",
    "BP",
    "BQ",
    "BR",
)

# 本当に空白のままでよい列
BLANK_IF_BLANK_TARGETS = (
    "AO",
    "BS",
    "BT",
    "BX",
    "BY",
)


DEFAULT_MAPPING = MappingConfig(
    always_copy_rules=(
        # 毎月更新の基本列
        CopyRule(source_col="F", target_col="A"),   # 新規案件 -> 新規･追加
        CopyRule(source_col="H", target_col="C"),   # 今月契約有無
        CopyRule(source_col="C", target_col="G"),   # 未収金額
        CopyRule(source_col="D", target_col="H"),   # 違約金未収額
        CopyRule(source_col="I", target_col="M"),   # 未収回数
        CopyRule(source_col="J", target_col="P"),   # 前月未収金額
        CopyRule(source_col="L", target_col="R"),   # 最初の未収
        CopyRule(source_col="M", target_col="S"),   # 最後の未収
        CopyRule(source_col="G", target_col="AR"),  # 先月契約有無 -> 先月ステータス

        # 常時転記に寄せた列
        CopyRule(source_col="A",  target_col="E"),   # 企業No
        CopyRule(source_col="B",  target_col="F"),   # 企業名
        CopyRule(source_col="N",  target_col="T"),   # 主商材
        CopyRule(source_col="O",  target_col="U"),   # 商材ステータス
        CopyRule(source_col="AJ", target_col="Z"),   # 重複項目
        CopyRule(source_col="AI", target_col="AA"),  # 重複
        CopyRule(source_col="S",  target_col="AB"),  # 責任者
        CopyRule(source_col="T",  target_col="AC"),  # 担当者
        CopyRule(source_col="V",  target_col="AK"),  # 口座番号
        CopyRule(source_col="W",  target_col="AT"),  # 請求方法
        CopyRule(source_col="AG", target_col="BA"),  # 非請求金額
        CopyRule(source_col="AH", target_col="BB"),  # 全額非請求フラグ
        CopyRule(source_col="Y",  target_col="BN"),  # 順番待ち有効
        CopyRule(source_col="Z",  target_col="BO"),  # 予約台帳有効
        CopyRule(source_col="AB", target_col="BP"),  # 食べログ有効
        CopyRule(source_col="AC", target_col="BQ"),  # LP有効
        CopyRule(source_col="AA", target_col="BR"),  # LINE有効

        # 追加で常時転記にしている列
        CopyRule(source_col="U",  target_col="D"),   # 対応方針
        CopyRule(source_col="P",  target_col="V"),   # 債権会社
        CopyRule(source_col="Q",  target_col="W"),   # パートナーコード
        CopyRule(source_col="R",  target_col="X"),   # パートナー名
        CopyRule(source_col="AE", target_col="BV"),  # 未収リスト除外
        CopyRule(source_col="AF", target_col="BW"),  # 除外理由
    ),
    conditional_copy_rules=(
        # いったん空にしている
    ),
    auto_number_rules=(
        AutoNumberRule(target_col="B"),
    ),
    fixed_value_rules=(
        # 常に固定値
        FixedValueRule(target_col="O", value="-"),

        # 値が入らなかったときだけ "-" を補完
        *(FixedValueRule(target_col=col, value="-", only_if_blank=True) for col in HYPHEN_IF_BLANK_TARGETS),

        # 値が入らなかったときだけ空白のまま維持
        *(FixedValueRule(target_col=col, value=None, only_if_blank=True) for col in BLANK_IF_BLANK_TARGETS),
    ),
    value_map_copy_rules=(
        ValueMapCopyRule(
            source_col="U",
            target_col="BU",
            value_map={
                "直接回収": 1,
                "弁護士委任": 2,
                "弁護士委任予定": 3,
                "全額非請求": 4,
                "売上取消": 5,
                "その他": 6,
                "第二次売却予定": 7,
                "法律事務所委任済": 8,
                "-": 9,
                # "合算請求予定": 要確認
            },
            default_value=None,
        ),
    ),
    formula_autofill_enabled=True,
)