# Project Vision

## Project Overview

Build a private personal mobile fitness tracking app for one user.

This app is not intended for public release, subscriptions, social features, or multi-user support.

The purpose of the app is to help me track and understand my personal fitness progress over time through workouts, nutrition, body measurements, progress photos, and analytics.

The app interface must be in English.

---

## Main Goal

Create a personal fitness operating system that helps me answer these questions:

- Am I getting stronger?
- Am I gaining muscle?
- Am I losing fat?
- Am I consistent with training?
- Am I hitting my calories and macros?
- Are my body measurements improving?
- What changed over the last week, month, 3 months, 6 months, and year?
- What should I adjust next?

---

## Core Focus Areas

The app should focus on:

1. Workout tracking
2. Nutrition tracking
3. Body measurements
4. Progress photos
5. Visual analytics
6. Long-term progress comparison

---

## What This App Should NOT Include

Do not build:

- Public user accounts
- Social feed
- Friends system
- Community features
- Subscriptions
- Payment system
- Marketing pages
- Public profiles
- Coach marketplace
- Complicated onboarding
- Multi-user administration
- Enterprise-level backend complexity

This is a private personal tool.

---

## Design Direction

The app should feel:

- Modern
- Clean
- Premium
- Minimal
- Fast
- Personal
- Data-driven
- Easy to use daily

The design should be simple but polished.

Avoid clutter. Avoid unnecessary screens. Avoid features that do not directly support personal fitness tracking.

---

## User Profile

The app is designed for one person who:

- Trains for hypertrophy
- Wants to improve body composition
- Wants to reduce belly and waist fat
- Wants to build upper-body muscle
- Prefers machines and cable exercises
- Tracks calories and macros
- Tracks protein, carbs, fat, and calories
- Wants periodic body measurements
- Wants progress photos
- Wants charts that clearly show improvement over time

---

## Key Product Philosophy

This app should not only store data. It should explain progress.

Raw numbers are not enough. Every statistic should help me understand what is happening.

Examples:

- "Your chest measurement increased by 2.5 cm over the last 8 weeks."
- "Your average weekly workout consistency is 86%."
- "Your protein target was achieved on 5 out of 7 days this week."
- "Your waist decreased while your weight stayed stable, which may indicate body recomposition."
- "Your pull volume is lower than your push volume this month."

The app should help me make better decisions.

---

## Primary Navigation

Use a simple bottom navigation structure:

1. Dashboard
2. Workouts
3. Nutrition
4. Measurements
5. Analytics

Keep navigation simple and consistent.

---

## Dashboard

The Dashboard should show a clear daily overview:

- Today's workout
- Workout completion status
- Calories consumed
- Protein consumed
- Remaining macros
- Current weight
- Latest body measurements summary
- Weekly consistency
- Quick actions

Quick actions:

- Start Workout
- Log Meal
- Add Weight
- Add Measurements
- Add Progress Photos

---

## Workout Module

Support:

- Workout programs
- Workout templates
- Exercise list
- Exercise history
- Sets
- Reps
- Weight
- RPE
- Rest timer
- Notes
- Workout duration
- Personal records
- Volume tracking
- Muscle group tracking

The app should make logging workouts fast.

When starting a workout, previous performance should be visible.

For each exercise, show:

- Last used weight
- Last reps
- Previous best
- Recent trend
- Target sets and reps

---

## Nutrition Module

Support daily tracking for:

- Calories
- Protein
- Carbohydrates
- Fat
- Fiber
- Water

The app should show:

- Daily targets
- Consumed amount
- Remaining amount
- Weekly average
- Monthly average
- Target adherence

Keep food logging simple.

Since this is a personal app, support custom repeated meals and quick meal templates.

Examples:

- Protein shake
- Chicken and rice
- Greek yogurt
- Eggs
- Peanut butter toast
- Restaurant meals

---

## Measurements Module

Support periodic tracking of:

- Weight
- Body fat %
- Neck
- Chest
- Waist
- Hips
- Left arm
- Right arm
- Left forearm
- Right forearm
- Left thigh
- Right thigh
- Left calf
- Right calf

The app should allow comparison between any two dates.

Show:

- Absolute change
- Percentage change
- Trend direction
- Best measurement
- Latest measurement

---

## Progress Photos

Support:

- Front photo
- Side photo
- Back photo

Allow:

- Date-based photo comparison
- Side-by-side comparison
- Before and after view
- Monthly progress review

Keep this section private and simple.

---

## Analytics Module

This is one of the most important parts of the app.

Analytics should include:

### Workout Analytics

- Total workouts
- Weekly workout consistency
- Monthly workout consistency
- Training frequency
- Total volume
- Volume by muscle group
- Volume trend over time
- Exercise PRs
- Estimated strength progress
- Missed workouts
- Most trained muscles
- Least trained muscles

### Nutrition Analytics

- Average daily calories
- Average daily protein
- Average daily carbs
- Average daily fat
- Macro adherence
- Calorie adherence
- Weekly nutrition trend
- Monthly nutrition trend

### Body Analytics

- Weight trend
- Waist trend
- Chest trend
- Arm trend
- Thigh trend
- Body fat trend
- Measurement changes over time
- Body recomposition indicators

### Progress Insights

The app should generate simple insight cards, such as:

- "Weight is stable, but waist is decreasing."
- "Protein intake has been below target for 3 days."
- "Workout consistency improved compared to last week."
- "Chest and arm measurements are trending upward."
- "Leg training volume is lower than upper body volume."

---

## Charts

Charts should be:

- Clean
- Easy to read
- Interactive if possible
- Not overloaded
- Useful at a glance

Support time ranges:

- 7 days
- 30 days
- 3 months
- 6 months
- 1 year
- All time

Use charts only when they help explain progress.

---

## Technical Direction

Build the app using modern mobile development practices.

Use:

- TypeScript
- Clean architecture
- Feature-based folder structure
- Reusable components
- Strong typing
- Local-first data storage if possible
- Simple backup/export support if practical

Prioritize maintainability.

Avoid unnecessary backend complexity unless truly needed.

---

## Engineering Principles

Follow these rules:

1. Keep files small and focused.
2. Avoid duplicated logic.
3. Separate UI from business logic.
4. Create reusable components.
5. Use clear naming.
6. Use consistent formatting.
7. Prefer simple solutions.
8. Avoid over-engineering.
9. Make future changes easy.
10. Keep the app fast.

---

## AI Development Behavior

When working on this project:

- Think like a senior mobile engineer.
- Think like a product designer.
- Think like a fitness coach.
- Think like a data analyst.

Before implementing any feature, consider:

- Is this useful for daily use?
- Is this simple enough?
- Does this help me understand progress?
- Can this be maintained easily?
- Is there a better UX approach?

Do not add public-app features unless specifically requested.

This is a private personal fitness tracker.
