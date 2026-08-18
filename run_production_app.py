"""Launch the PCS production desktop application."""

from production_runtime import apply_production_environment


apply_production_environment()

from run_app import main  # noqa: E402


if __name__ == "__main__":
    main()
