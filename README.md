# Bing Wallpaperer

Handy Powershell script to help update Windows desktop wallpaper directly from Bing's Image of the Day web service.

I like Bing's daily images, and I like having them changing daily as my desktop wallpaper. So some time ago I created this handy little script, which is setup as a scheduled task to run daily. It's been tweaked and improved over time - generally just as an excuse to dabble in some Powershell!

Having dusted it off and given it a bit of tidy up recently, I thought I'd also get it up into a repository to make it easier to manage and so it can be shared with others if anyone finds it useful or wants to build on it.
 

## Examples

Just run as-is and it will run with defaults...

`.\BingWallpaperer.ps1`

It will default to the `en-GB` market, but you can specify a different market like so...

`.\BingWallpaperer.ps1 -mkt en-US`

By default the script uses your current screen resolution (on the primary monitor) - not all image sizes are supported, but the script does try all the most common if the specified one is not available.  You can specify the image size in a `width`x`height` format like so...

`.\BingWallpaperer.ps1 -size 1366x768`

You can grab any image from the last 7 days. It defaults to the latest image (today's image), which is represented as index `0` and the maximum index is `7`. So for example to use the image from yesterday, you can specify the index like so...

`.\BingWallpaperer.ps1 -idx 1`


## Things to do (one day, maybe)

- Improve image sizes ... include portrait maybe.
- Add option to setup schedule task via script.