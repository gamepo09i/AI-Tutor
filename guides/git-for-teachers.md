# Git for Teachers

This is the simple version.

## The basic idea

Think of Git like save points for a folder.

- Your folder is the working copy.
- A commit is a named save point.
- GitHub is the online copy people can download from.
- Pushing uploads your latest commits to GitHub.

## Do I add files slowly or all at once?

Usually, add a useful batch all at once.

Good batches:

- "Add first version of slides"
- "Update class 2 setup instructions"
- "Add game starter files"
- "Fix typos in curriculum"

Avoid huge mystery batches like:

- "stuff"
- "new things"
- "final final version"

## Daily workflow

Run this when you want to see what changed:

```bash
git status
```

Add everything you changed:

```bash
git add .
```

Create a save point:

```bash
git commit -m "Add vibe coding slides"
```

Upload it to GitHub:

```bash
git push
```

## When students or other teachers need the newest version

They can download the repo from GitHub, or if they already cloned it, they can run:

```bash
git pull
```

## Recommended folder pattern

- `slides/` for presentation files
- `docs/` for curriculum plans and written material
- `examples/` for starter projects students can edit
- `guides/` for setup instructions

## A good rhythm for this course

Commit after each meaningful teaching update:

1. Add or revise slides.
2. Test that the slides open.
3. Commit.
4. Push to GitHub.

That way, people always get a version that makes sense.
