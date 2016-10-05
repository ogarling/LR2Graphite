# LR2Graphite
[LoadRunner](<http://www8.hp.com/nl/nl/software-solutions/loadrunner-load-testing/>) response time metrics export tool to [Graphite](https://graphite.readthedocs.io/en/latest/).

## Why?
When using LoadRunner in combination with another load test tool like Gatling, JMeter, etc. it would be nice to unify performance test metrics like response time statistics into a single datasource like Graphite.  
Works excellent in combination with [Targets-io](https://github.com/dmoll1974/targets-io) tool.

## What?
By using an [AutoIt](https://www.autoitscript.com/site/) script or executable the LoadRunner analysis MS Access database is being queried and the results are exported into Graphite.
Metrics are aggregated per 10 seconds (configurable) and exported to Graphite in "buckets". Currently not supported but when using [Statsd](https://github.com/etsy/statsd) the aggregation can be skipped.

## Which metrics can be exported?
At the moment the following response time metrics can be exported per transaction:  
- average
- minimum
- maximum
- (99th) percentile (configurable [0-100])

## How to use?
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
Jenkins (but preferably use LRlauncher.exe if possible):
```
LR2Graphite.exe <path to Jenkins job workspace> <Graphite host> <Graphite port> <timezone offset (hours)>  
e.g. LR2Graphite.exe "%WORKSPACE%" 123.123.123.123 2003 0  
```


## Prerequisites
- Analysed LoadRunner testrun with a MDB file. Note: do not select OUTPUT.MDB  
- Graphite instance to export metrics to

OS: Windows  
AutoIt for editing code and debugging.  

## Add-ons: support tooling
- LR2Graphite_modify_test_start_time.exe  
- LRlauncher.exe

### LR2Graphite_modify_test_start_time  
Simple tool for debugging purposes to manipulate the start time of a LoadRunner test to n hours in the past.  
```
Usage: LR2Graphite_modify_test_start_time.exe <number of hours in the past>
```
### Targets-io LRlauncher
Tool for creating a performance test pipeline:
- starting a LoadRunner test (command prompt CLI or via Jenkins)
- sending start, keepalive and end test events to Targets-io
- running LR2Graphite to export metrics to Graphite (so implicitely to Targets-io)
- validate requirements of test result via Targets-io  

Targets-io LRlauncher is described in more detail in this [Wiki](https://github.com/ogarling/LR2Graphite/wiki).
