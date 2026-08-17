"""Launch the isolated PCS production integration desktop application."""

from integration_runtime import apply_integration_environment


apply_integration_environment()

from run_app import main  # noqa: E402


if __name__ == "__main__":
    main()
