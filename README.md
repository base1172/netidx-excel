This is an Excel COM add-in that allows you subscribe to real time data from Netidx. You can also write data to a Netidx container.

Say you have published some data, maybe from the shell publisher, and you want to show it to your boss, but "that commandline thing" or "that linux thing" is not something your boss does.
```bash
> netidx publisher --bind 192.168.0.0/24 --spn publish/ryouko.ryu-oh.org@RYU-OH.ORG
/test/foo|array|[42, -42]
/test/bar|i32|42
/test/baz|i32|-42
```

If you install the add-in on you bosses' machine, then he can pull that data into a cell in an Excel spreadsheet, and you can stop worrying about how to present results to pointy haired types and get your job done.

![Example](example.PNG)

The key point is if a value in netidx updates, Excel will update almost immediatly, and your bosses pivot table will update too. csv exports are one step closer to dieing!
# Syntax

```
=NetGet(PATH)
```

`PATH` can of course be a ref, or another formula, it's Excel, your boss knows Excel ... right?

Alternatively, you can subscribe using Excel's built-in `=RTD()` function:
```
=RTD("NetidxRTD",,PATH)
```
There should be no material difference in performance between subscribing to data using `=NetGet()` versus using `=RTD()`.

# Writing

You can write to a Netidx path using the `=NetSet()` function. This requires that the publisher at that path supports write requests. Generally, `=NetSet()` should be used for writing to a Netidx container.
```
=NetSet(PATH,VALUE,[TYPE])
```
If the `TYPE` parameter is omitted, the add-in will try to choose an appropriate type for the value you are publishing. Alternatively, you can explicitly specify one of the below types:
* "auto"
* "f64"
* "i64"
* "null"
* "time"
* "string"
* "bool"

# Performance 

Even if you subscribe to a lot of data, or you subscribe to data that updates quickly, Excel should remain responsive because RTDs are throttled, and all the netidx processing is happening on a background thread pool. For example here Excel is maxing out my wifi network by subscribing to the stress publisher, however it remains completely responsive. It's actually pulling in 2 million updates per second, and that's limited by the network, not the cpu.

![Performance](perf.PNG)

# Building

There are pre built binaries [here](https://github.com/estokes/netidx-excel/releases/tag/0.1.6)

But you want to build it you need a windows machine (obviously) probably windows 10 or above, but 8 might work. You need to install rust, the easiest way is to use rustup. You need to install git. Once you have those two things installed,

```bash
> git clone https://github.com/estokes/netidx-excel
> cd netidx-excel
> cargo build --release
```

The dll should be built in `target/release/netidx_excel.dll`

# Installing

To install you need to decide where you want the add-in to live, it really doesn't matter where, but you should rename it to have an XLL extension (I use `netidx_excel.xll`). Then open Excel `Options | Add-ins`, select `Excel Add-ins`, and click `Go...`. Then `Browse` to the location where you have saved the XLL.

## 32 bit office on 64 bit windows

If you are running the 32 bit version of office, maybe because you have limited ram, then you will need to also install the netidx_excel32.dll, and you will need to run regsvr32 on that as well, just like the above. If you are building from source you will need to install the target `i686-pc-windows-msvc` and build the 32 bit dll with that target, e.g. `cargo build --target i686-pc-windows-msvc --release`, and then the dll will be in `target/i686-pc-windows-msvc/release` instead of `target/release`.

# Limitations

- No publish support; there's no real reason other than time, it's perfectly possible
- No resolver list support; once again, time, no real problems with this

# Other

I programmed on windows for a WHOLE month so that you NEVER have too. Because trust me, you NEVER want to. But if you are curious about the dreams I had during that month, read this before bed [Inside COM+](https://www.thrysoee.dk/InsideCOM+/ch05c.htm). Really, don't do it. COM must have seemed like a good idea to someone at some point in history, right? Developers! Developers! Developers! ... Developers! I mean, GObject seems totally great now, really.
