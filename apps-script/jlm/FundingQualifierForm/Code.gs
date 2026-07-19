function myFunction() {
  function createFundingQualifierForm() {
      // Create the form
        var form = FormApp.create('Funding Qualifier (2 minutes)');

          // Form settings
            form.setDescription(
                'Answer these so I can route you to the right next step. ' +
                    'This is not a loan application. No guarantees—this is readiness + packaging.'
                      );
                        form.setCollectEmail(true);
                          form.setLimitOneResponsePerUser(false);
                            form.setShowLinkToRespondAgain(false);

                              // Q1: Full name
                                form.addTextItem()
                                    .setTitle('Full name')
                                        .setRequired(true);

                                          // Q2: Best phone number
       
                                            form.addTextItem()
                                                .setTitle('Best phone number')
                                                    .setRequired(true);

                                                      // Q3: Business name + state
                                                        form.addTextItem()
                                                            .setTitle('Business name + state')
                                                                .setRequired(true);

                                                                  // Q4: Average monthly revenue (last 3 months)
                                                                    form.addMultipleChoiceItem()
                                                                        .setTitle('Average monthly revenue (last 3 months)')
                                                                            .setChoices([
                                                                                  form.createChoice('Under $5k'),
                                                                                        form.createChoice('$5k–$10k'),
                                                                                              form.createChoice('$10k–$25k'),
                                                                                                    form.createChoice('$25k+')
                                                                                                        ])
                                                                                                            .setRequired(true);

                                                                                                              // Q5: Time in business
                                                                                                                form.addMultipleChoiceItem()
                                                                                                                    .setTitle('Time in business')
                                                                                                                        .setChoices([
                                                                                                                              form.createChoice('<3 months'),
                                                                                                                                    form.createChoice('3–6 months'),
                                                                                                                                          form.createChoice('6–12 months'),
                                                                                                                                                form.createChoice('1+ year')
                                                                                                                                                    ])
                                                                                                                                                        .setRequired(true);

                                                                                                                                                          // Q6: Bank statements available
                                                                                                                                                            form.addMultipleChoiceItem()
                                                                                                                                                                .setTitle('Do you have 3–6 months business bank statements available (PDF)?')
                                                                                                                                                                    .setChoices([
                                                                                                                                                                          form.createChoice('Yes'),
                                                                                                                                                                                form.createChoice('No')
                                                                                                                                                                                    ])
                                                                                                                                                                                        .setRequired(true);

                                                                                                                                                                                          // Q7: NSFs/overdrafts
                                                                                                                                                                                            form.addMultipleChoiceItem()
                                                                                                                                                                                                .setTitle('Any NSFs/overdrafts or negative balance days in the last 90 days?')
                                                                                                                                                                                                    .setChoices([
                                                                                                                                                                                                          form.createChoice('None'),
                                                                                                                                                                                                                form.createChoice('1–2'),
                                                                                                                                                                                                                      form.createChoice('3+')
                                                                                                                                                                                                                          ])
                                                                                                                                                                                                                              .setRequired(true);

                                                                                                                                                                                                                                // Q8: Funding goal + use of funds
                                                                                                                                                                                                                                  form.addParagraphTextItem()
                                                                                                                                                                                                                                      .setTitle('Funding goal + what it’s for (one sentence)')
                                                                                                                                                                                                                                          .setRequired(true);

                                                                                                                                                                                                                                            // Log URLs
                                                                                                                                                                                                                                              Logger.log('Form edit URL: ' + form.getEditUrl());
                                                                                                                                                                                                                                                Logger.log('Form live URL: ' + form.getPublishedUrl());

                                                                                                                                                                                                                                                  return {
                                                                                                                                                                                                                                                      editUrl: form.getEditUrl(),
                                                                                                                                                                                                                                                          liveUrl: form.getPublishedUrl()
                                                                                                                                                                                                                                                            };
                                                                                                                                                                                                                                                            }

  }

