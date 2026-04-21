import argparse

from assessment_runner_core import (
    read_template_controls,
    write_blank_assessment_workbook,
    write_rubric_workbook,
    write_template_workbook,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate the detailed assessment template and blank workbook."
    )
    parser.add_argument("--template", required=True, help="Source template workbook.")
    parser.add_argument(
        "--template_output",
        default="security_assessment_template.xlsx",
        help="Path for the generated blank assessment template workbook.",
    )
    parser.add_argument(
        "--rubric_output",
        default="security_assessment_rubric.xlsx",
        help="Path for the generated standalone rubric workbook.",
    )
    parser.add_argument(
        "--report_template_output",
        default="assessment_report_template.xlsx",
        help="Path for the generated unfilled assessment report workbook.",
    )
    args = parser.parse_args()

    controls = read_template_controls(args.template)
    write_template_workbook(controls, args.template_output)
    write_rubric_workbook(args.rubric_output)
    write_blank_assessment_workbook(controls, args.report_template_output)

    print(f"Wrote template workbook: {args.template_output}")
    print(f"Wrote rubric workbook: {args.rubric_output}")
    print(f"Wrote assessment report template workbook: {args.report_template_output}")


if __name__ == "__main__":
    main()
