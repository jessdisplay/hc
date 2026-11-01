/* What the long checklist adds to the Sheet's Everything tab, which files each document once.
   Taken from the 2023 Avaana checklist, where a document sits under every standard it serves. This repo is
   public, so documents are numbers only: the page finds each title and link in Harmony's own data. */
window.HC_STANDARDS = {
  // documents that also serve a standard, by the number on Harmony's register
  also: {
    'Person-centred supports': [5, 6, 73, 76],
    'Individual values and beliefs': [6, 73, 76],
    'Independence and Informed Choice': [73, 76],
    'Risk Management': [66, 85],
    'Incident Management': [38, 39, 40],
    'Human Resource Management': [11],
    'Continuity of Supports': [45, 72],
    'Support Planning': [9, 72],
    'Service Agreements with Participants': [72],
    'Responsive Support Provision': [33, 72, 78],
    'Transitions to or from a provider': [72],
    'Safe environment': [18, 72],
    'Management of Medication': [45],
    'Regulated Restrictive Practices': [123],
    'Supporting the Assessment and Development of Behaviour Support Plans': [123],
    'Behaviour Support Plan Implementation': [123],
    'Monitoring and Reporting the Use of Regulated Restrictive Practices': [123],
    'Behaviour Support Plan Review': [123],
    'Reportable Incidents involving the Use of a Restrictive Practice': [123],
    'Interim Behaviour Support Plans': [123],
    'Specialised Support Coordination': [138],
    'Conflict of Interest': [9, 138]
  },
  // filed as evidence on the Sheet, listed as documents in the checklist
  documents: [14, 15],
  // what the 2023 checklist asks for where the Sheet's Evidence column has nothing
  evidence: {
    'Governance and Operational Management': ['Company constitution and organisational documents', 'Employment or contractor agreements with each worker'],
    'Information Management': ['Evidence of accountancy engagement', 'Accounting software', 'Cloud storage or a practice management system'],
    'Complex Bowel Care': ['A Registered Nurse on the team, with their qualifications'],
    'Enteral (Naso-Gastric Tube - Jejunum or Duodenum) Feeding and Management': ['A Registered Nurse on the team, with their qualifications'],
    'Behaviour Support in the NDIS': ['Experience implementing positive behaviour supports', 'Qualifications'],
    'Specialised Support Coordination': ['Relevant qualifications as a Psychologist, Occupational Therapist, Social Worker or Mental Health Nurse', 'CPD portfolio showing relevant training in the field']
  },
  // a Core Practice Standard the Sheet does not have, word for word from the NDIS Commission's standards
  added: {
    after: 'Service Agreements with Participants',
    name: 'Responsive Support Provision',
    outcome: 'Each participant accesses responsive, timely, competent and appropriate supports to meet their needs, desired outcomes and goals.\n(Practice Standards Rules, Schedule 1, clause 21; Quality Indicators Guidelines, section 22)',
    indicators: [
      'Supports are provided based on the least intrusive options, in accordance with contemporary evidence-informed practices that meet participant needs and help achieve desired outcomes.',
      'For each participant (with their consent or direction and as agreed in their service agreement) links are developed and maintained by the provider through collaboration with other providers, including health care and allied health providers, to share their information, manage risks to them and meet their needs.',
      'Reasonable efforts are made to involve the participant in selecting their workers, including the preferred gender of workers providing personal care supports.',
      'Where a participant has specific needs which require monitoring and/or daily support, workers are appropriately trained and understand the participant’s needs and preferences.'
    ],
    also: [],
    items: []
  }
};
