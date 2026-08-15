from __future__ import annotations

from dataclasses import dataclass
from itertools import product

from .checksum import calculate_check_digit, validate_td3_checksums
from .normalize import MRZ_CHARSET, normalize_text


@dataclass(frozen=True)
class CorrectionCandidate:
    line_1: str
    line_2: str
    corrections: tuple[str, ...] = ()
    data_edits: int = 0
    check_digit_matches: int = 0

    @property
    def edit_count(self) -> int:
        return len(self.corrections)

    @property
    def priority(self) -> tuple[int, int, int]:
        return (self.check_digit_matches, -self.edit_count, -self.data_edits)


_DIGIT_ALTERNATIVES = {
    "O": ("0",),
    "Q": ("0",),
    "D": ("0",),
    "I": ("1",),
    "L": ("1",),
    "Z": ("2",),
    "B": ("8",),
    "S": ("5",),
    "G": ("6",),
}

_ALPHA_ALTERNATIVES = {
    "0": ("O",),
    "1": ("I",),
    "2": ("Z",),
    "8": ("B",),
    "5": ("S",),
    "6": ("G",),
}

_ALNUM_ALTERNATIVES = {
    **_DIGIT_ALTERNATIVES,
    **_ALPHA_ALTERNATIVES,
}

_DATA_RANGES = tuple(list(range(0, 9)) + list(range(13, 19)) + list(range(21, 27)) + list(range(28, 42)))
_CHECK_DIGIT_POSITIONS = (9, 19, 27, 42, 43)


def _options(character: str, field_type: str) -> tuple[str, ...]:
    if field_type == "numeric":
        alternatives = _DIGIT_ALTERNATIVES.get(character, ())
    elif field_type == "alpha":
        alternatives = _ALPHA_ALTERNATIVES.get(character, ())
    else:
        alternatives = _ALNUM_ALTERNATIVES.get(character, ())
    values = [character]
    for alternative in alternatives:
        if alternative not in values:
            values.append(alternative)
    return tuple(values)


def _field_type(index: int) -> str:
    if 13 <= index <= 18 or 21 <= index <= 26:
        return "numeric"
    if 10 <= index <= 12:
        return "alpha"
    return "alnum"


def _expected_field_digits(line_2: str) -> dict[int, str]:
    return {
        9: calculate_check_digit(line_2[0:9]),
        19: calculate_check_digit(line_2[13:19]),
        27: calculate_check_digit(line_2[21:27]),
        42: calculate_check_digit(line_2[28:42]),
    }


def _original_check_match_count(line_2: str) -> int:
    expected = _expected_field_digits(line_2)
    composite_input = line_2[0:10] + line_2[13:20] + line_2[21:43]
    expected[43] = calculate_check_digit(composite_input)
    return sum(line_2[index] == digit for index, digit in expected.items())


def _name_correction(line_1: str) -> tuple[str, list[str]]:
    output = list(line_1)
    corrections = []
    for index in range(5, 44):
        replacement = _ALPHA_ALTERNATIVES.get(output[index])
        if replacement:
            corrections.append(f"line1[{index}]:{output[index]}->{replacement[0]}")
            output[index] = replacement[0]
    return "".join(output), corrections


def _beam_data_variants(
    line_2: str,
    max_candidates: int,
    max_data_edits: int,
) -> list[tuple[str, tuple[str, ...], int]]:
    beam: list[tuple[str, tuple[str, ...], int]] = [(line_2, (), 0)]
    for index in _DATA_RANGES:
        next_beam: list[tuple[str, tuple[str, ...], int]] = []
        for current, corrections, edits in beam:
            field_type = _field_type(index)
            for option in _options(current[index], field_type):
                if option not in MRZ_CHARSET:
                    continue
                if option == current[index]:
                    next_beam.append((current, corrections, edits))
                elif edits < max_data_edits:
                    updated = current[:index] + option + current[index + 1 :]
                    next_beam.append(
                        (
                            updated,
                            corrections + (f"line2[{index}]:{current[index]}->{option}",),
                            edits + 1,
                        )
                    )
        dedup: dict[str, tuple[str, tuple[str, ...], int]] = {}
        for candidate in next_beam:
            dedup.setdefault(candidate[0], candidate)
        beam = sorted(dedup.values(), key=lambda item: (item[2], len(item[1])))[:max_candidates]
        if not beam:
            break
    return beam


def _repair_check_digits(
    data_line_2: str,
    base_corrections: tuple[str, ...],
    data_edits: int,
    max_edits: int,
    original_matches: int,
) -> list[CorrectionCandidate]:
    expected_fields = _expected_field_digits(data_line_2)
    field_options: list[tuple[str, ...]] = []
    field_labels: list[tuple[int, str]] = []
    for index in (9, 19, 27, 42):
        observed = data_line_2[index]
        expected = expected_fields[index]
        field_options.append((observed,) if observed == expected else (observed, expected))
        field_labels.append((index, expected))

    output: list[CorrectionCandidate] = []
    for selected_fields in product(*field_options):
        field_edits = sum(selected != data_line_2[index] for selected, (index, _) in zip(selected_fields, field_labels))
        if data_edits + field_edits > max_edits:
            continue
        line_with_fields = list(data_line_2)
        corrections = list(base_corrections)
        for selected, (index, expected) in zip(selected_fields, field_labels):
            if selected != data_line_2[index]:
                corrections.append(f"line2[{index}]:{data_line_2[index]}->{selected}")
            line_with_fields[index] = selected

        line_with_fields = "".join(line_with_fields)
        composite_expected = calculate_check_digit(line_with_fields[0:10] + line_with_fields[13:20] + line_with_fields[21:43])
        composite_options = (line_with_fields[43],)
        if line_with_fields[43] != composite_expected and data_edits + field_edits < max_edits:
            composite_options = (line_with_fields[43], composite_expected)

        for composite in composite_options:
            total_edits = data_edits + field_edits + (composite != line_with_fields[43])
            if total_edits > max_edits:
                continue
            final_line = line_with_fields[:43] + composite
            if not validate_td3_checksums(final_line).all_valid:
                continue
            final_corrections = list(corrections)
            if composite != line_with_fields[43]:
                final_corrections.append(f"line2[43]:{line_with_fields[43]}->{composite}")
            output.append(
                CorrectionCandidate(
                    line_1="",
                    line_2=final_line,
                    corrections=tuple(final_corrections),
                    data_edits=data_edits,
                    check_digit_matches=original_matches,
                )
            )
    return output


def generate_correction_candidates(
    line_1: str,
    line_2: str,
    max_candidates: int = 256,
    max_edits: int = 3,
) -> list[CorrectionCandidate]:
    """Generate bounded candidates and retain only checksum-valid repairs."""
    line_1 = normalize_text(line_1)
    line_2 = normalize_text(line_2)
    if len(line_1) != 44 or len(line_2) != 44:
        return []
    if any(character not in MRZ_CHARSET for character in line_1 + line_2):
        return []

    corrected_line_1, name_corrections = _name_correction(line_1)
    candidates: list[CorrectionCandidate] = []
    for data_line_2, data_corrections, data_edits in _beam_data_variants(
        line_2,
        max_candidates=max_candidates,
        max_data_edits=max_edits,
    ):
        repaired = _repair_check_digits(
            data_line_2=data_line_2,
            base_corrections=tuple(name_corrections) + data_corrections,
            data_edits=data_edits,
            max_edits=max_edits,
            original_matches=_original_check_match_count(data_line_2),
        )
        for candidate in repaired:
            candidates.append(
                CorrectionCandidate(
                    line_1=corrected_line_1,
                    line_2=candidate.line_2,
                    corrections=candidate.corrections,
                    data_edits=candidate.data_edits,
                    check_digit_matches=candidate.check_digit_matches,
                )
            )

    candidates.sort(key=lambda item: item.priority, reverse=True)
    unique: dict[tuple[str, str], CorrectionCandidate] = {}
    for candidate in candidates:
        unique.setdefault((candidate.line_1, candidate.line_2), candidate)
    return list(unique.values())[:max_candidates]
