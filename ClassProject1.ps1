﻿Connect-AzAccount
Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"

############################
### Define These Variables
############################
$resourceGroupName = 'ClassResourceGroup'
$location = 'centralus'
$vm1 = 'ELKServer'
$protocol = 'TCP'
$healthcheckprotocol = 'http'
$insideport = 80
$outsideport = 80
$healthcheckpath = "/"
$adminuser  = 'AzureAdmin'

####################
## Global Config
####################
$me = get-azcontext | select -ExpandProperty Account | select -ExpandProperty Id
$homeip = Invoke-RestMethod jsonip.com | select -ExpandProperty ip

 Create an SSH Key in azure acceptable format
if (!(test-path "$profile\.ssh\id_rsa.pub")){
  ssh-keygen -C $me -t rsa -b 4096 -f $profile\.ssh\id_rsa -N '""'
}

#####################################
## Functions to build environment
#####################################

Function Create-AzEnvironment($resourceGroupName, $location) {
  $vnetName = 'ELK'
  $CIDRRange = '10.1.0.0/16'
  $subnetname = 'default'
  $subnet1 = '10.1.0.0/24'

  # Resource Group
  New-AzResourceGroup -Name $resourceGroupName `
    -Location $location

  # Subnet
  $Subnet = New-AzVirtualNetworkSubnetConfig `
    -Name $subnetname `
    -AddressPrefix $subnet1

  # Virtual Network
  $virtualNetwork = New-AzVirtualNetwork `
    -ResourceGroupName $resourceGroupName `
    -Location $location `
    -Name $vnetname `
    -AddressPrefix $CIDRRange `
    -Subnet $subnet

  return $virtualNetwork
}

Function Create-CustomVM ($VMName, $location, $resourceGroupName, $availabilitySetName, $VirtulNetwork, $adminuser){

  $securityGroupName = "$VMName-security-group"
  $NICName = "$VMName-NIC"
  $computerName = $VMName
  $PublicIPAddressName = "$VMName-public-ip"
  $vmsize = 'Standard_B1ms'
  $vmpublisher = 'Canonical'
  $vmoffer = 'UbuntuServer'
  $vmskus = '18.04-LTS'
  $AdminPassword = ConvertTo-SecureString "p4ssw0rd*" -AsPlainText -Force
  $Credential = New-Object System.Management.Automation.PSCredential ($AdminUser, $AdminPassword);

  # Security Group Config
  $rule1 = New-AzNetworkSecurityRuleConfig -Name "ssh-rule" `
    -Description "Allow SSH from Home IP" `
    -Access Allow `
    -Protocol Tcp `
    -Direction Inbound `
    -Priority 100 `
    -SourceAddressPrefix "$homeip" `
    -SourcePortRange * `
    -DestinationAddressPrefix * `
    -DestinationPortRange 22

     #Security Group
  $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName `
    -Location $location `
    -Name $securityGroupName `
    -SecurityRules $rule1

  # Network Interface
  $NIC = New-AzNetworkInterface -Name $NICName `
    -ResourceGroupName $ResourceGroupName `
    -Location $Location `
    -SubnetId $virtualNetwork.Subnets[0].Id

  # Public IP Address
  $PIP = New-AzPublicIpAddress -Name $PublicIPAddressName `
    -ResourceGroupName $resourceGroupName `
    -AllocationMethod Dynamic `
    -Location $location

  # Attach Public IP to NIC
  $nic | Set-AzNetworkInterfaceIpConfig -Name $($nic.IpConfigurations.name) `
    -PublicIPAddress $pip `
    -Subnet $subnet | Out-Null
  $nic | Set-AzNetworkInterface | out-null

  if ($availabilitySetName -ne $null){
    $availabilityset = Get-AzAvailabilitySet -name $availabilitySetName
    if (!($availabilityset)){
      $availabilityset = New-AzAvailabilitySet -Name "$availabilitySetName" `
        -ResourceGroupName $resourceGroupName `
        -Location $location `
        -PlatformFaultDomainCount 2 `
        -PlatformUpdateDomainCount 2 `
        -Sku Aligned
    }

    # Virtual Machine Config
    $VirtualMachine = New-AzVMConfig -VMName $VMName `
      -VMSize $VMSize `
      -AvailabilitySetId $availabilityset.id
  }
  else {
    # Virtual Machine Config
    $VirtualMachine = New-AzVMConfig -VMName $VMName `
      -VMSize $VMSize
  }



  # Operating System Config
  $VirtualMachine = Set-AzVMOperatingSystem -VM $VirtualMachine `
    -Linux -ComputerName $computerName `
    -Credential $Credential `
    -DisablePasswordAuthentication

  # Attach Network Interface
  $VirtualMachine = Add-AzVMNetworkInterface -VM $VirtualMachine `
    -Id $NIC.Id

  # Set Source Image
  $VirtualMachine = Set-AzVMSourceImage -VM $VirtualMachine `
    -PublisherName $vmpublisher `
    -Offer $vmoffer `
    -Skus $vmskus `
    -Version latest

  # Turn off Boot Diagnostics
  $VirtualMachine = Set-AzVMBootDiagnostic -VM $VirtualMachine `
    -Disable

  # Retrieve SSH Key data
  $sshPublicKey = Get-Content "$profile\.ssh\id_rsa.pub"

  # Add it to virtual machine config settings
  $VirtulMachine = Add-AzVMSshPublicKey -VM $virtualMachine `
    -KeyData $sshpublickey `
    -Path "/home/$AdminUser/.ssh/authorized_keys"

  # Create Virtual Machine
  $vm = New-AzVM -ResourceGroupName $ResourceGroupName `
    -Location $Location `
    -VM $VirtualMachine `
    -Verbose `
    -AsJob

  if ($availabilitySetName -ne $null) {
    return @{username = $AdminUser ; availabilityset = $availabilityset }
  }
  else {
    return @{username = $AdminUser }
  }
}
# Custom Function to run Remote Bash Commands
Function Run-AZRemoteBash($Script, $VMName, $ResourceGroupName, [switch]$AsJob){
  $myscript = @"
$script
"@
  set-content -Path "./temp-file.sh" -Value $myscript
  $scriptpath = Get-Item -Path "./temp-file.sh"
  if($AsJob){
    Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName `
      -VMName $vmName `
      -CommandId 'RunShellScript' `
      -ScriptPath "$($scriptpath.FullName)" `
      -AsJob
  }else{
    Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName `
      -VMName $vmName `
      -CommandId 'RunShellScript' `
      -ScriptPath "$($scriptpath.FullName)"
  }
  remove-item "./temp-file.sh"
}

Function Create-CustomLoadbalancer ($Name, $Protocol, $HealthCheckProtocol, $InsidePort, $OutsidePort, $HealthCheckPath, $ResourceGroupName, $Location) {
  $lbpublicip = New-AzPublicIpAddress -ResourceGroupName $resourceGroupName -Name "$Name-PublicIp" -Location $location -AllocationMethod "Dynamic"
  $frontend = New-AzLoadBalancerFrontendIpConfig -Name "$Name-FrontEnd" -PublicIpAddress $lbpublicip
  $backendAddressPool = New-AzLoadBalancerBackendAddressPoolConfig -Name "$Name-backendpool"
  $probe = New-AzLoadBalancerProbeConfig -Name "$Name-Probe" -Protocol $healthcheckprotocol -Port $insideport -IntervalInSeconds 15 -ProbeCount 2 -RequestPath $healthcheckpath
  # $inboundNatRule1 = New-AzLoadBalancerInboundNatRuleConfig -Name "$Name-NATRule" -FrontendIpConfiguration $frontend -Protocol $protocol -FrontendPort $outsideport -BackendPort $outsideport -IdleTimeoutInMinutes 15 -EnableFloatingIP
  $lbrule = New-AzLoadBalancerRuleConfig -Name "$Name-rule" -FrontendIpConfiguration $frontend -BackendAddressPool $backendAddressPool -Probe $probe -Protocol $protocol -FrontendPort $outsideport -BackendPort $insideport -IdleTimeoutInMinutes 15 -EnableFloatingIP -LoadDistribution SourceIPProtocol
  $lb = New-AzLoadBalancer -Name $Name -ResourceGroupName $resourceGroupName -Location $location -FrontendIpConfiguration $frontend -BackendAddressPool $backendAddressPool -Probe $probe -InboundNatRule $inboundNatRule1 -LoadBalancingRule $lbrule
  return $(Get-AzLoadBalancer -Name $Name -ResourceGroupName $resourceGroupName)
}

Function Add-NicToLoadBalancer($NetworkInterface, $BackendPoolConfig) {
  $NetworkInterface | Set-AzNetworkInterfaceIpConfig -Name $($NetworkInterface.IpConfigurations.name) -LoadBalancerBackendAddressPool $BackendPoolConfig | Out-Null
  $NetworkInterface | Set-AzNetworkInterface -AsJob | Out-Null
}




########################
## Build Our Stuff
########################
Write-Host 'Building the Network...' -ForegroundColor Yellow
#Build the network
$virtualnetwork = Create-AzEnvironment -resourceGroupName $resourceGroupName -location $location
Write-Host 'Network Build Complete.' -ForegroundColor Green


Write-Host 'Starting VM Build Jobs...' -ForegroundColor Yellow
# Building the VMs
$results = @{}
$results[$vm1] = Create-CustomVM -VMName $vm1 -location $location -resourceGroupName $resourceGroupName -VirtulNetwork $virtualnetwork -adminuser $adminuser
Write-Host 'Job 1 Started...' -ForegroundColor Yellow
#$results[$vm2] = Create-CustomVM -VMName $vm2 -location $location -resourceGroupName $resourceGroupName -VirtulNetwork $virtualnetwork -adminuser $adminuser -availabilitySetName 'RedTeamAS'
#Write-Host 'Job 2 Started...' -ForegroundColor Yellow
#$results[$vm3] = Create-CustomVM -VMName $vm3 -location $location -resourceGroupName $resourceGroupName -VirtulNetwork $virtualnetwork -adminuser $adminuser -availabilitySetName 'RedTeamAS'
#Write-Host 'Job 3 Started...' -ForegroundColor Yellow

# Wait for completion of vms.
Write-Host "Waiting for VM Builds to Complete..." -ForegroundColor Yellow
$i = 0
while($i -lt 1){
  $i = 0
  $stats = Get-AzVM
  $status1 = $stats | where {$_.name -eq $vm1}
  if($status1.ProvisioningState -eq 'Succeeded'){$i++}
  #$status2 = $stats | where {$_.name -eq $vm1}
  #if($status2.ProvisioningState -eq 'Succeeded'){$i++}
  #$status3 = $stats | where {$_.name -eq $vm1}
 # if($status3.ProvisioningState -eq 'Succeeded'){$i++}
  start-sleep 5
}
Write-Host "Finished Building VMs." -ForegroundColor Green

$vm1ip = Get-AzNetworkInterface -Name "$VM1-NIC" -ResourceGroupName $resourceGroupName | Get-AzNetworkInterfaceIpConfig | select -ExpandProperty PrivateIpAddress
#$vm2ip = Get-AzNetworkInterface -Name "$VM2-NIC" -ResourceGroupName $resourceGroupName | Get-AzNetworkInterfaceIpConfig | select -ExpandProperty PrivateIpAddress
#$vm3ip = Get-AzNetworkInterface -Name "$VM3-NIC" -ResourceGroupName $resourceGroupName | Get-AzNetworkInterfaceIpConfig | select -ExpandProperty PrivateIpAddress

######################
## Configure VM1
######################
Write-Host "Configuring VMs and Ansible" -ForegroundColor Yellow
Write-Host "Configuring VM 1" -ForegroundColor Yellow
$installscript = @'
apt-get update && apt install docker.io -y && service docker start && docker pull cyberxsecurity/ansible && exit
'@
Run-AZRemoteBash -Script $installscript `
 -VMName $vm1 `
 -ResourceGroupName $resourceGroupName
Write-Host "VM1 Configuration Complete." -ForegroundColor Green
######################
## Configure Ansible
######################
$sshPrivateKey = Get-Content "$profile\.ssh\id_rsa"
$sshPrivateKey = $sshPrivateKey -replace("`r`n", "`n") | out-string
$sshPublicKey = Get-Content "$profile\.ssh\id_rsa.pub"

$playbook = @"
---
- name: Configure Elk VM with Docker
  hosts: elk
  remote_user: AzureAdmin
  become: true
  tasks:
    # Use apt module
  - name: Install docker.io
    become: true
    apt:
      force_apt_get: yes
      name: docker.io
      state: present

    # Use apt module
  - name: Install python3-pip
    apt:
      force_apt_get: yes
      name: python3-pip
      state: present

    # Use pip module
  - name: Install Docker module
    pip:
      name: docker
      state: present

    # Use command module
  - name: Increase virtual memory
    command: sysctl -w vm.max_map_count=262144

    # Use shell module
  - name: Increase virtual memory on automatically on VM restart
    shell: echo "vm.max_map_count=262144" >> /etc/sysctl.conf

    # Use docker_container module
  - name: download and launch a docker elk container
    docker_container:
      name: elk
      image: sebp/elk
      state: started
      restart_policy: always
    # Please list the ports that ELK runs on
      published_ports:
        - "5601:5601"
        - "9200:9200"
        - "5044:5044" 

"@


$dockerscript = "ansible-playbook /etc/ansible/dvwa-playbook.yml --user $adminuser -vvv"
$createpersistence = @"
echo "[webservers]" > /home/$AdminUser/hosts && \
#echo "$vm2ip" >> /home/$AdminUser/hosts && \
#echo "$vm3ip" >> /home/$AdminUser/hosts && \
echo "[default]" > /home/$AdminUser/ansible.cfg && \
echo "remote_user = $AdminUser" >> /home/$AdminUser/ansible.cfg && \
echo "$sshPrivateKey" > /home/$AdminUser/.ssh/id_rsa && \
echo "$sshPublicKey" > /home/$AdminUser/.ssh/id_rsa.pub && \
sudo chmod 600 /home/$AdminUser/.ssh/* && \
sudo chown $($AdminUser + ":" +$AdminUser) /home/$AdminUser/.ssh/* && \
echo "$playbook" > /home/$AdminUser/dvwa-playbook.yml && \
mkdir -p /home/$AdminUser/roles && \
sudo docker container run --rm -it --mount type=bind,source="/home/$AdminUser/",target=/etc/ansible/ --mount type=bind,source="/home/$AdminUser/.ssh/",target=/root/.ssh/ cyberxsecurity/ansible $dockerscript
exit
"@

Run-AZRemoteBash -Script $createpersistence -VMName $vm1 -resourcegroupname $resourcegroupname

Write-Host "Completed VMs and Ansible" -ForegroundColor Green

#########################
## Create LoadBalancer
#########################
#Write-Host "Building Loadbalancer..." -ForegroundColor Yellow

#$loadbalancer = Create-CustomLoadbalancer `
#  -Name $lbname `
#  -Protocol $protocol `
#  -InsidePort $insideport `
#  -OutsidePort $outsideport `
#  -HealthCheckProtocol $healthcheckprotocol `
#  -HealthCheckPath $healthcheckpath `
#  -Location $location `
#  -ResourceGroupName $resourceGroupName


##############################
## Attach VMs to LoadBalancer
##############################
#$vms = Get-AZVm
#$VMsToLoadbalance = $vms | where { $_.name -like 'DVWA-*' }
#foreach ($VM in $VMsToLoadbalance){
#  $nic = Get-AzNetworkInterface -ResourceId $vm.NetworkProfile.NetworkInterfaces.id
#  Add-NicToLoadBalancer -NetworkInterface $nic -BackendPoolConfig $loadbalancer.BackendAddressPools | out-null
#}

#Write-Host "Loadbalancer Complete." -ForegroundColor Green

Write-Host "Build Complete." -ForegroundColor Green

$vms | % {
  $results[$_.Name]['publicip'] = $(Get-AzPublicIpAddress -Name "$($_.name)-public-ip" | Select-Object -ExpandProperty IpAddress)
  $results[$_.Name]['username'] = 'AzureAdmin'
}



##############################
## Finish
##############################
Write-Host "##############################`n## Results                       `n##############################" -ForegroundColor Blue
Write-Host "VM1: [$vm1][$vm1ip] - $adminuser@$($results[$vm1]['publicip'])" -ForegroundColor Green
#Write-Host "VM2: [$vm2][$vm2ip] - $adminuser@$($results[$vm2]['publicip'])" -ForegroundColor Green
#Write-Host "VM3: [$vm3][$vm2ip] - $adminuser@$($results[$vm3]['publicip'])" -ForegroundColor Green