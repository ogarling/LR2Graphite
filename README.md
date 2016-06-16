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
Standalone, manually or included in a pipeline. For example: [LoadRunner integration in Jenkins](https://wiki.jenkins-ci.org/display/JENKINS/HP+Application+Automation+Tools)   

Compiled executable can be run with (pipeline) or without command line options (manually/interactive)  
Manually (interactive):
```
LR2Graphite.exe  
```
Manually (command line):
```
LR2Graphite.exe <path to LR mdb> <Graphite host> <Graphite port> <timezone offset (hours)>  
```
Jenkins:
```
LR2Graphite.exe <path to Jenkins job workspace> <Graphite host> <Graphite port> <timezone offset (hours)>  
```


### Prerequisites
Analysed LoadRunner testrun with an MDB file. Note: do not select OUTPUT.MDB  
OS: Windows  
AutoIt for editing code and debugging.  
