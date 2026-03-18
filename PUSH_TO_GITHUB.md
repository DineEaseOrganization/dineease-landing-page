# How to Publish to GitHub Pages

## Step 1 – Save your 4 images

Copy your 4 app screenshots into the `images/` folder with these names:
- `images/slide1.png`  → DineEase splash / hero
- `images/slide2.png`  → Find screen
- `images/slide3.png`  → Book screen
- `images/slide4.png`  → Enjoy screen

## Step 2 – Create the GitHub repo

Go to https://github.com/organizations/DineEaseOrganization/repositories/new and create a new **public** repo named `dineease-landing-page`.

Leave it empty (no README, no .gitignore).

## Step 3 – Push from your terminal

Open a terminal, navigate to this folder, then run:

```bash
git add .
git commit -m "Initial DineEase landing page"
git remote add origin https://github.com/DineEaseOrganization/dineease-landing-page.git
git push -u origin main
```

## Step 4 – Enable GitHub Pages

1. Go to the repo on GitHub → **Settings** → **Pages**
2. Under **Source**, select **Deploy from a branch**
3. Choose branch: `main`, folder: `/ (root)`
4. Click **Save**

Your site will be live at:
👉 **https://dineeasiorganization.github.io/dineease-landing-page/**

(GitHub Pages can take 1–2 minutes to publish after the first push.)
