# After Effects MOGRT Compatibility Guide

This guide shows how to build an After Effects MOGRT that behaves well with Sub Creator.

## Goal

Sub Creator can:
- replace the subtitle text
- switch between animation modes when the template exposes the expected controls
- update many exposed controls from Premiere

Sub Creator does not create the word or line animation for you.
That animation must already exist in the After Effects template.

## Recommended control names in Essential Graphics

Use these names if you want the best compatibility.

### Required text control

Expose one text control with one of these names:
- `Text`
- `Source Text`

`Text` is the safest choice.

### Required animation controls for word or line modes

Expose these controls:
- `Animation`
- `Highlight Based On`

Recommended setup:
- `Animation` = `Checkbox Control`
- `Highlight Based On` = `Dropdown Menu Control`

Recommended dropdown order:
1. `Words`
2. `Lines`

Sub Creator expects that order when it switches between `Per word` and `Per line`.

## Optional controls

These are useful when your template supports them:
- `Font Size`
- `Characters Per Line`
- `Max Lines`

If your template already reacts to those controls, Sub Creator can reuse them.

## Recommended template structure

A simple structure is usually the most reliable.

Suggested setup:
- one control layer, for example `SC_CTRL`
- one main text layer, for example `TEXT_MASTER`
- optional duplicate text layers for separate word and line animation behaviors

Example:
- `TEXT_MASTER` for the base text and style
- `TEXT_WORD` for word-based animation
- `TEXT_LINE` for line-based animation

## How to build the animation logic

### Base text

Use one main text layer as the source of truth.
Expose that text in Essential Graphics as `Text`.

### Word mode

Create the real word-based animation in After Effects.
Typical options:
- text animators
- range selectors based on words
- opacity, scale, position, fill, or highlight changes per word

Drive its visibility or behavior from:
- `Animation`
- `Highlight Based On`

### Line mode

Create a separate line-based animation.
Typical options:
- line reveal
- line highlight
- line-by-line opacity or scale changes

Again, drive it from:
- `Animation`
- `Highlight Based On`

### No animation

When `Animation` is off, show a stable non-animated version of the text.

## Essential Graphics checklist

Before exporting the MOGRT, check this:

- only one main exposed text control
- `Animation` checkbox exposed
- `Highlight Based On` dropdown exposed
- dropdown order is exactly `Words`, then `Lines`
- optional controls use clear names such as `Font Size`, `Characters Per Line`, `Max Lines`
- the template still looks correct in Premiere without editing inside After Effects

## Good practices

- keep control names in English for the exposed controls above
- expose controls only once when possible
- avoid duplicate groups of identical controls unless they are really needed
- make sure the text layer accepts line breaks cleanly
- test the template in Premiere before distributing it

## Working with fonts and style

For After Effects-authored MOGRTs, Sub Creator can often read the current font information when the template exposes it correctly.

Still, the most practical workflow is often:
1. change the font in Premiere's `Properties` panel
2. read the selection in Sub Creator
3. use `Apply changes` to copy that style to the other selected clips when available

## What to avoid

Avoid relying on:
- undocumented control names
- unusual dropdown ordering for `Highlight Based On`
- duplicate exposed text controls with the same purpose
- complex template logic that only works when edited back in After Effects

## Summary

For the best Sub Creator compatibility, expose at least:
- `Text`
- `Animation`
- `Highlight Based On`

And make sure the actual animation behavior already exists in the After Effects template.
