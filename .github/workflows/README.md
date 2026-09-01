# GitHub Actions are temporarily paused

The workflow definitions are preserved in `.github/workflows-disabled/` so
GitHub does not discover or run them while the repository is being improved.

To re-enable every workflow, move the YAML files back into this directory:

```sh
git mv .github/workflows-disabled/*.yml .github/workflows/
git commit -m "ci: re-enable GitHub Actions"
```
