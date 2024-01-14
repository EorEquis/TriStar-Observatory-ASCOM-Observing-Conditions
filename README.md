# TriStar Observatory ASCOM Observing Conditions

An ASCOM observing conditions driver.

---

## Requirements

* A weather file accessible by HTTP
* The file must be formatted as so
```
{
  "LastWrite": "2024-01-14 17:13:49",
  "Temp": -0.62,
  "Hum": 48.1,
  "DewPoint": -10.23,
  "Pres": 30.2,
  "WSp": 11.15,
  "WGust": 15.52,
  "WDir": 0.0,
  "RTot": 0.0,
  "LightSen": 801,
  "SkyTemp": -18.94,
  "IRAmb": 7.0,
  "RSen": 496,
  "CloudCondition": 1,
  "DaylightCondition": 2,
  "RainCondition": 1,
  "WindCondition": 2,
  "RSenD": 1,
  "UnsafeWarning": "0000",
  "Alert": 0
}
```

## Why isn't this working for me?

* It's all pretty specific to my needs and setup.  This will *absolutely not* work for you out of the box.  It's here as a starting point to give you ideas.

## I can't figure out _____.

* I'm happy to try to help.  Just email me.  Please understand this isn't a job or a product or anything...I have a life and a job and all that jazz, so I'll get to it when I can.

## None of this works, my observatory exploded because your code sucks.

* Sucks to be you.  You downloaded someone else's code, free of charge, and trusted your observatory to it.  You get what you get.
* If you're not comfortable with placing such trust in a whole pile of shitty arduino and C# code, making your own tweaks, wiring up your own electronics, and DIY risks in general, this probably isn't the project for you.
* In short, I bear no responsibility for any outcome of your use of this code.  Use at your own risk!  
