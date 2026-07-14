# Formulas

Every calculation the app performs, transcribed from `index.html`. Names in
parentheses are the functions where each lives.

## Estimated one-rep max (`est1RM`)

Epley formula, rounded to the nearest integer. Returns `0` for non-positive
weight or reps.

```
1RM = round( weight × (1 + reps / 30) )
```

## Set volume

Volume of a single set is `weight × reps`. **Best volume** for an exercise is the
largest single-set volume among its working sets (`computePRs`).

## Workout volume (`workoutVolume`)

Sum of `weight × reps` over all **working** (non-warm-up) sets in the workout.

```
workoutVolume = Σ (weight × reps)   over working sets
```

Weekly volume (Analytics) sums `workoutVolume` over the workouts in each week.

## Total sets (`totalSets`)

Count of **all** sets across a workout's exercises, including warm-ups. (Note:
"working sets this week" in Analytics instead counts only non-warm-up sets.)

## Personal record check (`isPR`)

A set qualifies as a weight PR when:

```
set is not a warm-up
AND set.weight > 0
AND set.weight > max(weight of all working sets for this exercise
                     in workouts dated before this workout's date)
```

## Nutrition totals (`calcTotals`)

Element-wise sum of every meal's macros for the day:

```
calories = Σ meal.calories
protein  = Σ meal.protein
carbs    = Σ meal.carbs
fat      = Σ meal.fat
```

## Calorie ring (`renderHome`)

```
calorieRingFraction = min(1, eaten / dailyCalories)
caloriesRemaining   = max(0, dailyCalories − eaten)
```

## Macro bars (`setBar`)

```
barPercent = min(100, goal > 0 ? (value / goal × 100) : 0)
```

## Progress ring geometry (`setRing`, `updateTimerDisplay`)

An SVG ring is filled by adjusting `stroke-dashoffset` against a fixed
circumference:

```
strokeDashoffset = circumference × (1 − fraction)
```

Circumferences used: `326.7` (calorie ring), `502.7` (rest-timer ring,
`TIMER_CIRC`).

## Weight-goal progress (`renderHome`)

Given current weight `gw`, target `tw`, and all recorded weights:

```
losing        = tw < gw
effectiveSW   = losing ? max(allWeights, startWeight)
                       : min(allWeights, startWeight)
total         = |effectiveSW − tw|
progress      = total > 0 ? (effectiveSW − gw) / (effectiveSW − tw) : 1
percent       = clamp( round(progress × 100), 0, 100 )
kgToGo        = |gw − tw|
goalReached   = |gw − tw| < 0.5
```

## Training streak (`renderHome`)

Starting from today and walking backward one day at a time, increment the streak
for each consecutive day that has at least one workout; stop at the first gap.

## Week boundaries

The current week starts on **Monday**:

```
weekStartOffset = (date.getDay() + 6) % 7
```

The same Monday-based offset is used for the home weekly strip, weekly analytics
bucketing (`weekStartOf`, `lastNWeeks`), and the calendar grid layout.

## Muscle frequency (Analytics)

Working sets are counted per muscle group over the **last 30 days**, then sorted
descending, to show training frequency by muscle
(`getExerciseGroup` maps an exercise to its group).

## Estimated-1RM trend (Analytics)

For a selected exercise, each workout contributes one point equal to the best
`est1RM` among that workout's working sets for the exercise, plotted over time.
