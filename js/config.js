const API = 'https://script.google.com/macros/s/AKfycbx7qLbW1aELpEny4m7pwt8b8AQniNFdNxae8Qz_38hipvx2OX0S6hIotl2iunJfGXKi7g/exec';
const EDITOR_PASSWORD = 'capplan2026';
const DEFAULT_THR = {cluster_cpu:80,cluster_ram:90,storage_core:70,storage_support:80};
const MONTHS = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];

let resources=[], history=[], projections=[], thresholds={...DEFAULT_THR};
let isEditor = false;
let currentTab = '';
let selProjRes=null, selProjMetric=null, histFilterId=null;
let detailView='list', drillCharts={}, overviewChart=null, growthCharts={};
