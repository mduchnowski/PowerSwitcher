# PowerSwitcher
Control the power relay from Digital Loggers ProSwitch

# The Power Relay 
DLI LPC9

# Useful Resources

https://www.digital-loggers.com/rest.html

Make note of the following powerful curl command that is shard on this page

Change the state of several outlets as simultaneously as possible. (Requires firmware 1.10.11+)
curl --digest -u admin:1234 -H "X-CSRF: x"  -H "Content-type: application/json"  --data-binary "[[[0,false],[1,true],[2,false],[5,true],[7,true]]]" "http://192.168.0.100/restapi/relay/set_outlet_transient_states/" 


The firmware must be updated to include the patch that allows a delay value of zero to be set on the relay.
