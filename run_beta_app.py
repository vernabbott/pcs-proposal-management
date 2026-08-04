"""Launch the isolated PCS beta desktop application."""

from beta_runtime import apply_beta_environment


apply_beta_environment()

from run_app import main  # noqa: E402


if __name__ == "__main__":
    main()
