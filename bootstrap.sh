#!/bin/bash
sudo apt-get -y update
sudo apt-get -y upgrade
sudo apt-get -y install git
sudo apt-get -y install build-essential zlib1g-dev libreadline5-dev libssl-dev libxml2-dev libxslt-dev unzip screen
( curl http://rvm.beginrescueend.com/releases/rvm-install-latest ) | bash
echo '[[ -s "$HOME/.rvm/scripts/rvm" ]] && . "$HOME/.rvm/scripts/rvm"' > ~/.profile
source ~/.rvm/scripts/rvm
rvm install 1.9.2
rvm --default use 1.9.2
gem install rubyfromexcel
