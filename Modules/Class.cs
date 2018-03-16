using System;
using System.Collections.Generic;

public class User
{
    public string id {get;set;}
    public string name {get;set;}
    public List<Pool> Pools {get;set;}
}
public class Pool 
{
    public string Phase {get;set;}
    public string Platoon {get;set;}
    public List<Order> Orders {get;set;}
}
public class Order
{
    public string Character{get;set;}
    public int PlatoonNum{get;set;}
}