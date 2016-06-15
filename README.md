# LR2Graphite
[LoadRunner](<http://www8.hp.com/nl/nl/software-solutions/loadrunner-load-testing/>) response time metrics export tool to [Graphite](https://graphite.readthedocs.io/en/latest/).

### Why?
When using LoadRunner in combination with another load test tool like Gatling, JMeter, etc. it would be nice to unify performance test metrics like response time statistics into a single datasource like Graphite.

### What?
By using an [AutoIt](https://www.autoitscript.com/site/) script or executable the LoadRunner analysis MS Access database is being queried and the results are exported into Graphite.
Metrics are aggregated per 10 seconds (configurable) and exported to Graphite in "buckets". Currently not supported but when using [Statsd](https://github.com/etsy/statsd) the aggregation can be skipped.

### Which metrics can be exported?
At the moment the following response time metrics can be exported per transaction:  
- average
- minimum
- maximum
- (99th) percentile (configurable [0-100])

### How to use?
Compiled executable can be run with or without command line options:
```
LR2Graphite.exe <path to LR mdb> <Graphite host> <Graphite port>
```

### Prerequisites
OS: Windows  
AutoIt for editing code and debugging.  
